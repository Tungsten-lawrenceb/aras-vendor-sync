<#
.SYNOPSIS
    Drain DatasheetFetchRequest rows in state=Queued.

.DESCRIPTION
    For each Queued row, look up the linked Manufacturer Part's Mfr PN,
    call Digi-Key ProductDetails, download the datasheet PDF, upload
    to the Aras vault via the 3-step transfer, attach the new File to
    the MP via `Manufacturer Part File`, and PATCH the request to
    state=Completed (or Failed with a clear last_error).

    Runs as a Windows Scheduled Task every 15 min by default.
    Companion to Refresh-VendorPricing.ps1 (weekly pricing refresh).
    Shares OAuth + endpoint helpers; the two scripts intentionally
    duplicate small amounts of code rather than introduce a shared
    module right now.

.PARAMETER ConfigPath
    Path to the JSON config file. Defaults to
    C:\ProgramData\AarasVendorSync\config.json.

.PARAMETER DryRun
    Report what would happen without writing.

.PARAMETER Limit
    Max rows processed in one run. Defaults to 20 (keeps each run
    short for the 15-min cadence).
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$env:ProgramData\AarasVendorSync\config.json",
    [switch]$DryRun,
    [int]$Limit = 20
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ----------------------------------------------------------------------------
# Config + logging (mirrors Refresh-VendorPricing.ps1)
# ----------------------------------------------------------------------------

if (-not (Test-Path $ConfigPath)) {
    throw "Config not found: $ConfigPath"
}
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

$logDir = $config.log.directory
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logPath = Join-Path $logDir ("{0:yyyy-MM-dd}-drain.log" -f (Get-Date))

function Write-Log {
    param(
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'SKIP')] [string]$Level,
        [string]$Message
    )
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1,-5} {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $logPath -Value $line
    Write-Host $line
}

Write-Log INFO "=== aras-datasheet-drain starting (DryRun=$DryRun Limit=$Limit) ==="

# ----------------------------------------------------------------------------
# Aras OAuth
# ----------------------------------------------------------------------------

function Get-ArasToken {
    param($ArasConfig)
    $url = "$($ArasConfig.url.TrimEnd('/'))/OAuthServer/connect/token"
    $resp = Invoke-RestMethod -Method Post -Uri $url -Body @{
        grant_type = 'password'; scope = 'Innovator'; client_id = 'IOMApp'
        username   = $ArasConfig.username; password = $ArasConfig.password_md5
        database   = $ArasConfig.database
    } -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 30
    if (-not $resp.access_token) { throw "Aras token response missing access_token" }
    return $resp.access_token
}

$arasToken = Get-ArasToken $config.aras
$arasHeaders = @{ Authorization = "Bearer $arasToken"; Accept = 'application/json' }
$arasBase = $config.aras.url.TrimEnd('/')
$odataBase = "$arasBase/server/odata"
$vaultBase = "$arasBase/vault/odata"
Write-Log OK "Aras token acquired"

# ----------------------------------------------------------------------------
# MPN_API_Config and Digi-Key OAuth
# ----------------------------------------------------------------------------

$mpn = Invoke-RestMethod -Uri "$odataBase/MPN_API_Config?`$select=vendor,client_id,client_secret,default_endpoint" `
    -Headers $arasHeaders -TimeoutSec 30
$dkCfg = $mpn.value | Where-Object { $_.vendor -eq 'DigiKey' } | Select-Object -First 1
if (-not $dkCfg) { Write-Log ERROR "No DigiKey row in MPN_API_Config"; exit 1 }
$dkEndpoint = ($dkCfg.default_endpoint -as [string])
if (-not $dkEndpoint) { $dkEndpoint = 'https://api.digikey.com' }
$dkEndpoint = $dkEndpoint.TrimEnd('/')

$dkTok = Invoke-RestMethod -Method Post -Uri "$dkEndpoint/v1/oauth2/token" `
    -ContentType 'application/x-www-form-urlencoded' `
    -Body @{ grant_type = 'client_credentials'; client_id = $dkCfg.client_id; client_secret = $dkCfg.client_secret } `
    -TimeoutSec 30
$dkToken = $dkTok.access_token
$dkHeaders = @{ Authorization = "Bearer $dkToken"; 'X-DIGIKEY-Client-Id' = $dkCfg.client_id; Accept = 'application/json' }
Write-Log OK "Digi-Key token acquired"

# Resolve service user's default vault id (REST API section 5.1)
$me = Invoke-RestMethod -Uri "$odataBase/User?`$filter=login_name eq '$($config.aras.username)'&`$select=id,default_vault&`$top=1" `
    -Headers $arasHeaders -TimeoutSec 30
$vaultId = $me.value[0].'default_vault@aras.id'
if (-not $vaultId) { Write-Log ERROR "Could not resolve default_vault for $($config.aras.username)"; exit 1 }
Write-Log INFO "Default vault: $vaultId"

# ----------------------------------------------------------------------------
# xxHash32 (decimal string) - vault chunk-integrity header
# ----------------------------------------------------------------------------

# PowerShell 5.1-compatible xxHash32 implementation. Returns the hash as
# an unsigned-32 integer; the vault wants it formatted as a decimal
# string. Mirrors the Python worker's xxhash.xxh32(data).intdigest().
function Get-XxHash32Decimal {
    param([byte[]]$Data)
    $P1 = [uint32]2654435761; $P2 = [uint32]2246822519
    $P3 = [uint32]3266489917; $P4 = [uint32]668265263; $P5 = [uint32]374761393
    $mask = [uint64]0xFFFFFFFF

    function _mul([uint32]$a, [uint32]$b) {
        return [uint32]((([uint64]$a * [uint64]$b) -band $script:mask))
    }
    function _rotl([uint32]$v, [int]$n) {
        return [uint32]((($v -shl $n) -bor ($v -shr (32 - $n))) -band $script:mask)
    }
    function _round([uint32]$acc, [uint32]$input) {
        $acc = [uint32](([uint64]$acc + ([uint64](_mul $input $script:P2))) -band $script:mask)
        return _mul (_rotl $acc 13) $script:P1
    }

    $script:mask = $mask
    $script:P1 = $P1; $script:P2 = $P2

    $len = $Data.Length
    $i = 0
    if ($len -ge 16) {
        $v1 = [uint32]((([uint64]$P1 + [uint64]$P2) -band $mask))
        $v2 = $P2
        $v3 = [uint32]0
        $v4 = [uint32](([uint64]0 - [uint64]$P1) -band $mask)
        while (($len - $i) -ge 16) {
            $v1 = _round $v1 ([BitConverter]::ToUInt32($Data, $i))
            $v2 = _round $v2 ([BitConverter]::ToUInt32($Data, $i + 4))
            $v3 = _round $v3 ([BitConverter]::ToUInt32($Data, $i + 8))
            $v4 = _round $v4 ([BitConverter]::ToUInt32($Data, $i + 12))
            $i += 16
        }
        $h = [uint32]((([uint64](_rotl $v1 1) + [uint64](_rotl $v2 7) + [uint64](_rotl $v3 12) + [uint64](_rotl $v4 18)) -band $mask))
    } else {
        $h = $P5
    }
    $h = [uint32](([uint64]$h + [uint64]$len) -band $mask)
    while (($len - $i) -ge 4) {
        $w = [BitConverter]::ToUInt32($Data, $i)
        $h = [uint32](([uint64]$h + [uint64](_mul $w $P3)) -band $mask)
        $h = _mul (_rotl $h 17) $P4
        $i += 4
    }
    while ($i -lt $len) {
        $h = [uint32](([uint64]$h + [uint64](_mul ([uint32]$Data[$i]) $P5)) -band $mask)
        $h = _mul (_rotl $h 11) $P1
        $i += 1
    }
    $h = [uint32]($h -bxor ($h -shr 15))
    $h = _mul $h $P2
    $h = [uint32]($h -bxor ($h -shr 13))
    $h = _mul $h $P3
    $h = [uint32]($h -bxor ($h -shr 16))
    return $h.ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

# ----------------------------------------------------------------------------
# Aras vault 3-step upload
# ----------------------------------------------------------------------------

function New-ArasId {
    return ([guid]::NewGuid().ToString('N')).ToUpperInvariant()
}

function Invoke-VaultBegin {
    $resp = Invoke-RestMethod -Method Post -Uri "$vaultBase/vault.BeginTransaction" `
        -Headers ($arasHeaders + @{ VAULTID = $vaultId }) `
        -ContentType 'application/octet-stream' -Body ([byte[]]@()) `
        -TimeoutSec 30
    return $resp.transactionId
}

function Invoke-VaultUploadChunk {
    param([string]$TxnId, [string]$FileId, [string]$Filename, [byte[]]$Chunk, [long]$Offset, [long]$Total)
    $cs = Get-XxHash32Decimal $Chunk
    $headers = $arasHeaders + @{
        VAULTID                            = $vaultId
        transactionid                      = $TxnId
        'Content-Disposition'              = "attachment; filename*=utf-8''$([uri]::EscapeDataString($Filename))"
        'Content-Range'                    = "bytes $Offset-$($Offset + $Chunk.Length - 1)/$Total"
        'Aras-Content-Range-Checksum'      = $cs
        'Aras-Content-Range-Checksum-Type' = 'xxHashAsUInt32AsDecimalString'
    }
    Invoke-RestMethod -Method Post -Uri "$vaultBase/vault.UploadFile?fileId=$FileId" `
        -Headers $headers -ContentType 'application/octet-stream' -Body $Chunk -TimeoutSec 60 | Out-Null
}

function Invoke-VaultCommitFile {
    # Commit the upload as a bare File item. Returns the new File id.
    param([string]$TxnId, [string]$FileId, [string]$Filename, [long]$Size)
    $boundary = "batch_$([guid]::NewGuid().ToString('N'))"
    # IIS-rooted sub-request path: /InnovatorServer/server/odata/File
    $parsed = [uri]"$odataBase/File"
    $subPath = $parsed.AbsolutePath
    $host = $parsed.Host
    $payload = @{
        '@odata.type' = '#File'
        id            = $FileId
        filename      = $Filename
        file_size     = $Size
    } | ConvertTo-Json -Compress
    $body = (
        "--$boundary`r`n" +
        "Content-Type: application/http`r`n" +
        "Content-Transfer-Encoding: binary`r`n" +
        "`r`n" +
        "POST $subPath HTTP/1.1`r`n" +
        "Host: $host`r`n" +
        "Content-Type: application/json`r`n" +
        "`r`n" +
        "$payload`r`n" +
        "--$boundary--`r`n"
    )
    $headers = $arasHeaders + @{
        VAULTID         = $vaultId
        transactionid   = $TxnId
        'OData-Version' = '4.0'
    }
    $resp = Invoke-WebRequest -Method Post -Uri "$vaultBase/vault.CommitTransaction" `
        -Headers $headers -ContentType "multipart/mixed; boundary=$boundary" `
        -Body $body -TimeoutSec 60 -UseBasicParsing
    # Aras returns 204 + Location: .../File('FILEID') OR 200 with no Location.
    # In either case the File id is the one we passed in.
    return $FileId
}

function Upload-FileToVault {
    param([byte[]]$Bytes, [string]$Filename)
    $txnId = Invoke-VaultBegin
    $fileId = New-ArasId
    $chunkSize = 16MB
    $offset = 0
    while ($offset -lt $Bytes.Length) {
        $end = [Math]::Min($offset + $chunkSize, $Bytes.Length)
        $chunk = New-Object byte[] ($end - $offset)
        [Array]::Copy($Bytes, $offset, $chunk, 0, $end - $offset)
        Invoke-VaultUploadChunk -TxnId $txnId -FileId $fileId -Filename $Filename `
            -Chunk $chunk -Offset $offset -Total $Bytes.Length
        $offset = $end
    }
    Invoke-VaultCommitFile -TxnId $txnId -FileId $fileId -Filename $Filename -Size $Bytes.Length | Out-Null
    return $fileId
}

# ----------------------------------------------------------------------------
# Digi-Key lookup (mirrors Refresh-VendorPricing.ps1)
# ----------------------------------------------------------------------------

function Get-DigiKeyDatasheetUrl {
    param([string]$MfrPn)
    $encoded = [System.Uri]::EscapeDataString($MfrPn)
    try {
        $resp = Invoke-RestMethod -Method Get -Uri "$dkEndpoint/products/v4/search/$encoded/productdetails" `
            -Headers $dkHeaders -TimeoutSec 30
        if ($resp.Product -and $resp.Product.DatasheetUrl) {
            return @{ url = $resp.Product.DatasheetUrl; dk_pn = $resp.Product.DigiKeyProductNumber }
        }
    } catch { }
    # Keyword fallback
    try {
        $kwBody = @{ Keywords = $MfrPn; Limit = 5; Offset = 0 } | ConvertTo-Json -Compress
        $kw = Invoke-RestMethod -Method Post -Uri "$dkEndpoint/products/v4/search/keyword" `
            -Headers ($dkHeaders + @{ 'Content-Type' = 'application/json' }) -Body $kwBody -TimeoutSec 30
    } catch { return $null }
    if (-not $kw.Products) { return $null }
    $targetNorm = ($MfrPn -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    $exact = $kw.Products | Where-Object {
        ($_.ManufacturerProductNumber -replace '[^A-Za-z0-9]', '').ToUpperInvariant() -eq $targetNorm
    }
    $best = $exact | Where-Object { $_.DatasheetUrl -and $_.DatasheetUrl.Contains('mm.digikey.com') } | Select-Object -First 1
    if (-not $best) { $best = $exact | Where-Object { $_.DatasheetUrl } | Select-Object -First 1 }
    if (-not $best) { $best = $kw.Products | Where-Object { $_.DatasheetUrl } | Select-Object -First 1 }
    if ($best) { return @{ url = $best.DatasheetUrl; dk_pn = $best.DigiKeyProductNumber } }
    return $null
}

function Normalize-Url {
    param([string]$Url)
    if ($Url -and $Url.StartsWith('//')) { return "https:$Url" }
    return $Url
}

function Download-Pdf {
    param([string]$Url)
    $url = Normalize-Url $Url
    $browserHeaders = @{
        'User-Agent'      = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36'
        Accept            = 'application/pdf,application/octet-stream,text/html;q=0.5,*/*;q=0.1'
        'Accept-Language' = 'en-US,en;q=0.9'
    }
    $resp = Invoke-WebRequest -Method Get -Uri $url -Headers $browserHeaders `
        -TimeoutSec 60 -MaximumRedirection 10 -UseBasicParsing
    $bytes = $resp.Content
    if (-not $bytes -or $bytes.Length -lt 8) { throw "empty response" }
    $head = ($bytes[0..7] -join ' ')
    $magic = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min(8, $bytes.Length))
    $magic = $magic.TrimStart("`0`r`n`t ", [char]0xFEFF, [char]0xEF, [char]0xBB, [char]0xBF)
    if (-not $magic.StartsWith('%PDF')) {
        throw "URL did not return a PDF (Content-Type=$($resp.Headers.'Content-Type'); first 8 bytes: $head)"
    }
    return $bytes
}

# ----------------------------------------------------------------------------
# Helpers for the queue loop
# ----------------------------------------------------------------------------

function Get-ManufacturerPartItemNumber {
    param([string]$MpId)
    $r = Invoke-RestMethod -Uri "$odataBase/Manufacturer%20Part('$MpId')?`$select=item_number" `
        -Headers $arasHeaders -TimeoutSec 30
    return $r.item_number
}

function Update-Request {
    param([string]$RowId, [hashtable]$Body)
    if ($DryRun) { return }
    # Filter out null values to keep the JSON clean.
    $clean = @{}
    foreach ($k in $Body.Keys) {
        $v = $Body[$k]
        if ($null -ne $v) { $clean[$k] = $v }
    }
    $json = $clean | ConvertTo-Json -Compress

    # PS 5.1's Invoke-RestMethod has been observed to corrupt PATCH bodies
    # to specific Aras endpoints (Aras complains about 0x00 chars in the
    # body even though our JSON is pure ASCII — likely a UTF-16 encoding
    # quirk in the cmdlet). curl.exe sends bytes verbatim and works.
    $url = "$odataBase/DatasheetFetchRequest('$RowId')"
    $tmp = New-TemporaryFile
    try {
        # Write the JSON as raw bytes (ASCII compatible UTF-8 without BOM).
        [System.IO.File]::WriteAllBytes(
            $tmp.FullName,
            [System.Text.UTF8Encoding]::new($false).GetBytes($json))
        $authHeader = "Authorization: Bearer $arasToken"
        $ctHeader = 'Content-Type: application/json'
        $code = curl.exe -s -o NUL -w '%{http_code}' -X PATCH $url `
            -H $authHeader -H $ctHeader `
            --data-binary "@$($tmp.FullName)"
        if ($code -notmatch '^2\d\d$') {
            # Re-run to capture the body for diagnostics
            $resp = curl.exe -s -X PATCH $url `
                -H $authHeader -H $ctHeader `
                --data-binary "@$($tmp.FullName)"
            throw "PATCH $url -> HTTP ${code}: $resp"
        }
    } finally {
        Remove-Item $tmp.FullName -ErrorAction SilentlyContinue
    }
}

function New-ManufacturerPartFile {
    param([string]$MpId, [string]$FileId)
    if ($DryRun) { return }
    # Idempotent
    $existing = Invoke-RestMethod `
        -Uri "$odataBase/Manufacturer%20Part%20File?`$filter=source_id eq '$MpId' and related_id eq '$FileId'&`$select=id&`$top=1" `
        -Headers $arasHeaders -TimeoutSec 30
    if ($existing.value -and $existing.value.Count -gt 0) { return $existing.value[0].id }
    $body = @{ source_id = $MpId; related_id = $FileId } | ConvertTo-Json -Compress
    $resp = Invoke-RestMethod -Method Post -Uri "$odataBase/Manufacturer%20Part%20File" `
        -Headers ($arasHeaders + @{ 'Content-Type' = 'application/json' }) `
        -Body $body -TimeoutSec 30
    return $resp.id
}

function Convert-MpnToSafeFilename {
    param([string]$Mpn)
    $cleaned = ([char[]]$Mpn | ForEach-Object {
        if ($_ -match '[A-Za-z0-9\-_.]') { $_ } else { '_' }
    }) -join ''
    return "$cleaned.pdf"
}

# ----------------------------------------------------------------------------
# Main loop
# ----------------------------------------------------------------------------

$queued = Invoke-RestMethod `
    -Uri "$odataBase/DatasheetFetchRequest?`$filter=request_state eq 'Queued'&`$select=id,manufacturer_part_id,vendor,attempt_count&`$top=$Limit&`$orderby=created_on" `
    -Headers $arasHeaders -TimeoutSec 30
Write-Log INFO "Queued rows to drain: $($queued.value.Count)"

$ok = 0; $failed = 0; $skipped = 0
$nowIso = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

foreach ($req in $queued.value) {
    $rowId = $req.id
    $mpId = $req.'manufacturer_part_id@aras.id'
    $mfrPn = $req.'manufacturer_part_id@aras.keyed_name'
    $attempts = [int]$req.attempt_count

    try {
        Update-Request -RowId $rowId -Body @{
            request_state = 'Running'
            attempt_count = ($attempts + 1)
        }
    } catch { Write-Log WARN "could not mark Running: $($_.Exception.Message)" }

    try {
        if (-not $mfrPn) {
            $mfrPn = Get-ManufacturerPartItemNumber -MpId $mpId
        }
        if (-not $mfrPn) { throw "no item_number on MP $mpId" }

        Write-Log INFO "Processing $mfrPn (req=$rowId)"
        $lookup = Get-DigiKeyDatasheetUrl -MfrPn $mfrPn
        if (-not $lookup -or -not $lookup.url) {
            Update-Request -RowId $rowId -Body @{ request_state = 'Failed'; last_error = 'no datasheet URL from Digi-Key' }
            Write-Log SKIP "$mfrPn - no DK datasheet"
            $skipped++; continue
        }
        $bytes = Download-Pdf -Url $lookup.url
        $filename = Convert-MpnToSafeFilename -Mpn $mfrPn
        $fileId = Upload-FileToVault -Bytes $bytes -Filename $filename
        New-ManufacturerPartFile -MpId $mpId -FileId $fileId | Out-Null
        Update-Request -RowId $rowId -Body @{
            request_state = 'Completed'
            file_id       = $fileId
            datasheet_url = $lookup.url
            completed_on  = $nowIso
            last_error    = $null
        }
        Write-Log OK "$mfrPn - attached $($bytes.Length) bytes (file=$fileId)"
        $ok++
    } catch {
        $msg = "$($_.Exception.GetType().Name): $($_.Exception.Message)"
        Update-Request -RowId $rowId -Body @{ request_state = 'Failed'; last_error = $msg }
        Write-Log ERROR "$mfrPn - $msg"
        $failed++
    }
}

Write-Log INFO "Done: ok=$ok failed=$failed skipped=$skipped of $($queued.value.Count)"
exit ([int]($failed -gt 0))
