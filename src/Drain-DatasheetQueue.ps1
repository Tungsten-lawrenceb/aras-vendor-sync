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

# xxHash32 via inline-compiled C#. PowerShell's bitwise semantics for
# uint32 are subtly broken (signed shifts, intermediate overflows, scope
# capture in scriptblocks) and a pure-PS implementation kept tripping
# on edge cases. C# handles unsigned arithmetic natively.
#
# Add-Type compiles on first invocation (~100ms) and is cached after.
if (-not ('AraSync.XxHash32' -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @'
namespace AraSync {
    using System;
    public static class XxHash32 {
        const uint P1 = 0x9E3779B1U;
        const uint P2 = 0x85EBCA77U;
        const uint P3 = 0xC2B2AE3DU;
        const uint P4 = 0x27D4EB2FU;
        const uint P5 = 0x165667B1U;
        static uint Round(uint acc, uint input) {
            unchecked {
                acc += input * P2;
                acc = (acc << 13) | (acc >> 19);
                acc *= P1;
                return acc;
            }
        }
        static uint RotL(uint v, int n) {
            return (v << n) | (v >> (32 - n));
        }
        public static uint Compute(byte[] data) {
            unchecked {
                int len = data.Length;
                int i = 0;
                uint h;
                if (len >= 16) {
                    uint v1 = P1 + P2;
                    uint v2 = P2;
                    uint v3 = 0;
                    uint v4 = 0U - P1;
                    while (len - i >= 16) {
                        v1 = Round(v1, BitConverter.ToUInt32(data, i));
                        v2 = Round(v2, BitConverter.ToUInt32(data, i + 4));
                        v3 = Round(v3, BitConverter.ToUInt32(data, i + 8));
                        v4 = Round(v4, BitConverter.ToUInt32(data, i + 12));
                        i += 16;
                    }
                    h = RotL(v1, 1) + RotL(v2, 7) + RotL(v3, 12) + RotL(v4, 18);
                } else {
                    h = P5;
                }
                h += (uint)len;
                while (len - i >= 4) {
                    h += BitConverter.ToUInt32(data, i) * P3;
                    h = RotL(h, 17) * P4;
                    i += 4;
                }
                while (i < len) {
                    h += (uint)data[i] * P5;
                    h = RotL(h, 11) * P1;
                    i += 1;
                }
                h ^= h >> 15;
                h *= P2;
                h ^= h >> 13;
                h *= P3;
                h ^= h >> 16;
                return h;
            }
        }
    }
}
'@
}

function Get-XxHash32Decimal {
    param([byte[]]$Data)
    return [AraSync.XxHash32]::Compute($Data).ToString(
        [System.Globalization.CultureInfo]::InvariantCulture)
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
    # PS's TrimStart wants a char[] (not a string + char args). Defensively
    # check the leading bytes for any BOM or whitespace before looking for
    # the %PDF magic.
    $magic = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min(16, $bytes.Length))
    $trim = [char[]](0x00, 0x09, 0x0A, 0x0D, 0x20, 0xFEFF, 0xEF, 0xBB, 0xBF)
    $magic = $magic.TrimStart($trim)
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
    # Strip XML-illegal control chars (notably 0x00) from string values.
    # Aras's OData layer XML-translates the JSON body internally and chokes
    # on any 0x00 byte; sanitize here so a Python-style error message
    # containing   in `last_error` doesn't crash the patch.
    $sanitized = @{}
    foreach ($k in $clean.Keys) {
        $v = $clean[$k]
        if ($v -is [string]) {
            $sanitized[$k] = ($v -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", '?')
        } else {
            $sanitized[$k] = $v
        }
    }
    $json = $sanitized | ConvertTo-Json -Compress

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
