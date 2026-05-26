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

# Shared per-vendor throttle (Mouser 30/min, DigiKey ~100/min defaults).
. "$PSScriptRoot\Vendor-Throttle.ps1"

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

# Mouser uses a single API key (in `client_secret`), no OAuth.
$msCfg = $mpn.value | Where-Object { $_.vendor -eq 'Mouser' } | Select-Object -First 1
$msEndpoint = $null
$msApiKey = $null
if ($msCfg -and $msCfg.client_secret) {
    $msApiKey = $msCfg.client_secret
    $msEndpoint = ($msCfg.default_endpoint -as [string])
    if (-not $msEndpoint) { $msEndpoint = 'https://api.mouser.com' }
    $msEndpoint = $msEndpoint.TrimEnd('/')
    Write-Log OK "Mouser API key loaded"
} else {
    Write-Log INFO "Mouser API key not configured (fallback disabled)"
}

# Known Vendor item ids (matches src/aras_mcp/importers/datasheets.py VENDOR_IDS).
# Used by Upsert-VendorPart as the source_id of the Vendor Part rel row.
$VendorIds = @{
    DigiKey = 'E94E47CFF55149C9B16BBD2939D861C9'
    Mouser  = '947C86F817C2454D9C3AE7062D3E0669'
}

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
    # Use Invoke-WebRequest so we can read the raw response body. The
    # transactionId comes back as JSON but Invoke-RestMethod has been
    # observed to leave it empty in PS 5.1 — likely a content-encoding
    # parse quirk. Parse manually.
    $resp = Invoke-WebRequest -Method Post -Uri "$vaultBase/vault.BeginTransaction" `
        -Headers ($arasHeaders + @{ VAULTID = $vaultId }) `
        -ContentType 'application/octet-stream' -Body ([byte[]]@()) `
        -TimeoutSec 30 -UseBasicParsing
    $body = $resp.Content
    if ($body -is [byte[]]) {
        # Direct byte path: decode as UTF-8 and strip the BOM if present.
        $body = [System.Text.Encoding]::UTF8.GetString($body)
    } else {
        # Invoke-WebRequest returns Content as a string already decoded —
        # but in Windows-1252 / Latin-1 if no charset was specified.
        # Re-round-trip through bytes to recover the true UTF-8 string.
        $latin1 = [System.Text.Encoding]::GetEncoding(28591)
        $body = [System.Text.Encoding]::UTF8.GetString($latin1.GetBytes($body))
    }
    # Now the BOM (if any) is the U+FEFF Zero-Width No-Break Space char.
    $body = $body.TrimStart([char[]](0xFEFF, 0x00, 0x09, 0x0A, 0x0D, 0x20))
    $j = $body | ConvertFrom-Json
    $txn = $j.transactionId
    if (-not $txn) { $txn = $j.transaction_id }
    if (-not $txn) { throw "vault.BeginTransaction returned no transactionId: $body" }
    return $txn
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
    # NOTE: `$host` is a read-only PS built-in (the host application).
    # Use a different variable name here.
    $hostName = $parsed.Host
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
        "Host: $hostName`r`n" +
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
    Write-Log INFO "vault: Begin (size=$($Bytes.Length))"
    $txnId = Invoke-VaultBegin
    Write-Log OK   "vault: txn=$txnId"
    $fileId = New-ArasId
    $chunkSize = 16MB
    $offset = 0
    while ($offset -lt $Bytes.Length) {
        $end = [Math]::Min($offset + $chunkSize, $Bytes.Length)
        $chunk = New-Object byte[] ($end - $offset)
        [Array]::Copy($Bytes, $offset, $chunk, 0, $end - $offset)
        Write-Log INFO "vault: UploadFile chunk offset=$offset len=$($chunk.Length)"
        Invoke-VaultUploadChunk -TxnId $txnId -FileId $fileId -Filename $Filename `
            -Chunk $chunk -Offset $offset -Total $Bytes.Length
        $offset = $end
    }
    Write-Log INFO "vault: CommitTransaction fileId=$fileId"
    Invoke-VaultCommitFile -TxnId $txnId -FileId $fileId -Filename $Filename -Size $Bytes.Length | Out-Null
    Write-Log OK   "vault: committed"
    return $fileId
}

# ----------------------------------------------------------------------------
# MPN variants — loose matching helper
# ----------------------------------------------------------------------------

# Yield up to N variants of an MPN ordered most-likely-first. Each
# variant is a string we'll try against DK and Mouser. The first
# variant is always the original MPN.
function Get-MpnVariants {
    param([string]$Mpn, [int]$Max = 6)
    # In-line dedup: collect into a list, skip duplicates as we go.
    # (Prior version used a nested helper that captured $script:seen,
    # which leaked state across calls and broke variant generation on
    # subsequent invocations within the same script.)
    $out = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $candidates = @()
    $candidates += $Mpn
    # 1) Drop a trailing reel/packaging code: comma-or-space + alnum tail
    #    e.g. "PZU16BL,315" -> "PZU16BL", "ABM8-272-T3" -> "ABM8-272"
    if ($Mpn -match '^(.+),([A-Za-z0-9]+)$') { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)-T\d?$')           { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)-TR$')             { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)#TRPBF$')          { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)#PBF$')            { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)TR$')              { $candidates += $Matches[1] }
    # 2) Compress whitespace and try collapsed form
    $collapsed = ($Mpn -replace '\s+', '')
    if ($collapsed -ne $Mpn)                 { $candidates += $collapsed }
    # 3) Drop trailing space + token (e.g. "W25Q128JVSIQ TR" -> "W25Q128JVSIQ")
    if ($Mpn -match '^(.+)\s+\S+$')          { $candidates += $Matches[1] }
    foreach ($c in $candidates) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        $t = $c.Trim()
        if ($seen.Add($t)) { $out.Add($t) }
        if ($out.Count -ge $Max) { break }
    }
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# Digi-Key lookup — returns datasheet URL AND pricing.
# Tries the exact MPN first, then a few loose variants. For each variant
# we attempt ProductDetails (fast path), then keyword search.
# ----------------------------------------------------------------------------

function _ExtractDkPricing {
    param($Product)
    # ProductDetails shape: Product.ProductVariations[].StandardPricing[]
    # Note: in V4 the top-level Product.DigiKeyProductNumber is EMPTY —
    # the real SKU is on each variation. So we also return dk_pn from
    # whichever variation gave us the best price.
    if (-not $Product) { return $null }
    $variations = $Product.ProductVariations
    if (-not $variations) { return $null }
    $bestPrice = $null
    $bestCurr  = $null
    $bestMoq   = $null
    $bestDkPn  = $null
    foreach ($v in $variations) {
        $standardPricing = $v.StandardPricing
        if (-not $standardPricing) { continue }
        # Find lowest break-quantity row with a non-null UnitPrice.
        $rows = @($standardPricing | Where-Object { $_.UnitPrice -ne $null })
        if (-not $rows) { continue }
        $row = $rows | Sort-Object BreakQuantity | Select-Object -First 1
        $up = [decimal]$row.UnitPrice
        $moq = $v.MinimumOrderQuantity
        if (-not $moq) { $moq = $row.BreakQuantity }
        if ($null -eq $bestPrice -or $up -lt $bestPrice) {
            $bestPrice = $up
            $bestCurr  = if ($Product.PricingCurrency) { $Product.PricingCurrency } else { 'USD' }
            $bestMoq   = $moq
            $bestDkPn  = $v.DigiKeyProductNumber
        }
    }
    if ($null -eq $bestPrice) { return $null }
    return @{
        unit_price    = $bestPrice
        currency      = $bestCurr
        min_order_qty = $bestMoq
        dk_pn         = $bestDkPn
    }
}

function _DkProductDetails {
    param([string]$Mpn)
    $encoded = [System.Uri]::EscapeDataString($Mpn)
    Wait-ApiInterval -Api 'DigiKey'
    try {
        $resp = Invoke-RestMethod -Method Get -Uri "$dkEndpoint/products/v4/search/$encoded/productdetails" `
            -Headers $dkHeaders -TimeoutSec 30
        return $resp.Product
    } catch { return $null }
}

function _DkKeywordSearch {
    param([string]$Mpn)
    Wait-ApiInterval -Api 'DigiKey'
    try {
        $kwBody = @{ Keywords = $Mpn; Limit = 10; Offset = 0 } | ConvertTo-Json -Compress
        $kw = Invoke-RestMethod -Method Post -Uri "$dkEndpoint/products/v4/search/keyword" `
            -Headers ($dkHeaders + @{ 'Content-Type' = 'application/json' }) -Body $kwBody -TimeoutSec 30
    } catch { return $null }
    if (-not $kw.Products) { return $null }
    $targetNorm = ($Mpn -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    # STRICT: only accept an exact alnum-normalized match against the
    # requested MPN. Earlier versions fell back to "first hit with
    # datasheet" / "first hit", which produced bogus Vendor Part rows
    # for non-distributor MPNs (e.g. Tungsten custom PNs, hobby brands).
    # If none of the keyword results actually equal the MPN we asked
    # for, return null and let the caller move on.
    $exact = @($kw.Products | Where-Object {
            ($_.ManufacturerProductNumber -replace '[^A-Za-z0-9]', '').ToUpperInvariant() -eq $targetNorm
        })
    if ($exact.Count -gt 0) { return $exact[0] }
    return $null
}

function Get-DigiKeyLookup {
    param([string]$MfrPn)
    $variants = Get-MpnVariants -Mpn $MfrPn
    foreach ($v in $variants) {
        $vNorm = ($v -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
        $prod = _DkProductDetails -Mpn $v
        # ProductDetails can return a "close" match that isn't actually
        # the MPN we asked for. Verify the returned MfrPN alnum-equals
        # our variant before trusting it. Same strictness used by the
        # keyword search.
        if ($prod) {
            $rNorm = ($prod.ManufacturerProductNumber -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
            if ($rNorm -ne $vNorm) { $prod = $null }
        }
        if (-not $prod -or -not $prod.DatasheetUrl) {
            # ProductDetails miss — try keyword search (also strict)
            $prod = _DkKeywordSearch -Mpn $v
        }
        if (-not $prod) { continue }
        $url = $prod.DatasheetUrl
        $pricing = _ExtractDkPricing -Product $prod
        # V4: top-level DigiKeyProductNumber is usually empty. Prefer the
        # SKU from the lowest-priced variation; fall back to any
        # variation's SKU; finally fall back to the empty top-level field.
        $dkPn = $null
        if ($pricing) { $dkPn = $pricing.dk_pn }
        if (-not $dkPn -and $prod.ProductVariations -and $prod.ProductVariations.Count -gt 0) {
            $dkPn = $prod.ProductVariations[0].DigiKeyProductNumber
        }
        if (-not $dkPn) { $dkPn = $prod.DigiKeyProductNumber }
        if (-not $url -and -not $pricing -and -not $dkPn) { continue }
        $out = @{
            url           = $url
            dk_pn         = $dkPn
            matched_mpn   = $prod.ManufacturerProductNumber
            used_variant  = $v
        }
        if ($pricing) {
            $out.unit_price    = $pricing.unit_price
            $out.currency      = $pricing.currency
            $out.min_order_qty = $pricing.min_order_qty
        }
        return $out
    }
    return $null
}

# ----------------------------------------------------------------------------
# Mouser lookup — single API key, partnumber search.
# Returns datasheet URL AND pricing (first non-null PriceBreaks entry).
# ----------------------------------------------------------------------------

# Mouser Availability is sometimes "In Stock 1234" or just "1234".
# Pull the leading integer; null on anything we can't parse.
function _MouserParseAvailability {
    param([string]$S)
    if ([string]::IsNullOrWhiteSpace($S)) { return $null }
    $m = [regex]::Match($S, '\d+')
    if (-not $m.Success) { return $null }
    $n = 0
    if ([int]::TryParse($m.Value, [ref]$n)) { return $n }
    return $null
}

function _MouserParseMoney {
    param([string]$S)
    if ([string]::IsNullOrWhiteSpace($S)) { return $null }
    $m = [regex]::Match($S, '[\d.,]+')
    if (-not $m.Success) { return $null }
    $raw = $m.Value
    # If both comma and period appear, the rightmost is the decimal sep.
    if ($raw.Contains(',') -and $raw.Contains('.')) {
        if ($raw.LastIndexOf(',') -gt $raw.LastIndexOf('.')) {
            $raw = $raw.Replace('.', '').Replace(',', '.')
        } else {
            $raw = $raw.Replace(',', '')
        }
    } elseif ($raw.Contains(',') -and -not $raw.Contains('.')) {
        $raw = $raw.Replace(',', '.')
    }
    $d = 0
    if ([decimal]::TryParse($raw, [System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        return $d
    }
    return $null
}

function _MouserSearch {
    param([string]$Mpn)
    if (-not $msApiKey) { return $null }
    $body = @{
        SearchByPartRequest = @{
            mouserPartNumber  = $Mpn
            partSearchOptions = 'string'
        }
    } | ConvertTo-Json -Compress -Depth 4
    Wait-ApiInterval -Api 'Mouser'
    try {
        $resp = Invoke-RestMethod -Method Post `
            -Uri "$msEndpoint/api/v1/search/partnumber?apiKey=$msApiKey" `
            -Headers @{ Accept = 'application/json' } `
            -ContentType 'application/json' `
            -Body $body -TimeoutSec 30
    } catch { return $null }
    if (-not $resp -or -not $resp.SearchResults) { return $null }
    if ($resp.Errors -and $resp.Errors.Count -gt 0) { return $null }
    $parts = $resp.SearchResults.Parts
    if (-not $parts) { return $null }
    return $parts
}

function Get-MouserLookup {
    param([string]$MfrPn)
    if (-not $msApiKey) { return $null }
    $variants = Get-MpnVariants -Mpn $MfrPn
    foreach ($v in $variants) {
        $parts = _MouserSearch -Mpn $v
        if (-not $parts) { continue }
        $targetNorm = ($v -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
        # STRICT: only accept an exact alnum-normalized match. Mouser
        # search returns substring matches (e.g. "23165" -> "517-2316-
        # 5211TB"), and the older "first hit" / "startswith" fallbacks
        # created bogus Vendor Part rows.
        $exact = @($parts | Where-Object {
                ($_.ManufacturerPartNumber -replace '[^A-Za-z0-9]', '').ToUpperInvariant() -eq $targetNorm
            })
        $chosen = $null
        # Prefer an exact-match hit that also has a datasheet
        foreach ($p in $exact) {
            if ($p.DataSheetUrl) { $chosen = $p; break }
        }
        # Otherwise any exact-match hit, even without a datasheet
        if (-not $chosen -and $exact.Count -gt 0) { $chosen = $exact[0] }
        if (-not $chosen) { continue }

        # Pull pricing from PriceBreaks: first non-null entry.
        $unitPrice = $null
        $currency  = $null
        $minOrder  = $null
        if ($chosen.PriceBreaks) {
            foreach ($pb in $chosen.PriceBreaks) {
                $up = _MouserParseMoney $pb.Price
                if ($null -ne $up) {
                    $unitPrice = $up
                    $currency  = $pb.Currency
                    break
                }
            }
        }
        $moq = 0
        if (-not [int]::TryParse(($chosen.Min -as [string]), [ref]$moq)) { $moq = $null }

        $url = $chosen.DataSheetUrl
        if (-not $url -and $null -eq $unitPrice) { continue }
        $availability = _MouserParseAvailability $chosen.Availability
        $productStatus = $chosen.LifecycleStatus
        $out = @{
            url           = $url
            mouser_pn     = $chosen.MouserPartNumber
            matched_mpn   = $chosen.ManufacturerPartNumber
            used_variant  = $v
        }
        if ($null -ne $unitPrice) {
            $out.unit_price = $unitPrice
            $out.currency   = $currency
        }
        if ($moq) { $out.min_order_qty = $moq }
        if ($null -ne $availability) { $out.availability = $availability }
        if (-not [string]::IsNullOrWhiteSpace($productStatus)) {
            $out.product_status = $productStatus
        }
        return $out
    }
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

# Returns $true iff this MP already has at least one Manufacturer Part File
# rel — i.e. a datasheet (or any file) is already attached. We use this to
# avoid re-downloading + re-attaching duplicate datasheets when the queue
# is re-run for the same MP (which has happened repeatedly during pricing
# backfills). Note: this is a coarser check than New-ManufacturerPartFile,
# which only deduplicates exact (mp, file) pairs — useless when the new
# upload generates a fresh File id even for the same PDF content.
function Test-MpHasDatasheet {
    param([string]$MpId)
    $resp = Invoke-RestMethod `
        -Uri "$odataBase/Manufacturer%20Part%20File?`$filter=source_id eq '$MpId'&`$select=id&`$top=1" `
        -Headers $arasHeaders -TimeoutSec 30
    return ($resp.value -and $resp.value.Count -gt 0)
}

function Convert-MpnToSafeFilename {
    param([string]$Mpn)
    $cleaned = ([char[]]$Mpn | ForEach-Object {
        if ($_ -match '[A-Za-z0-9\-_.]') { $_ } else { '_' }
    }) -join ''
    return "$cleaned.pdf"
}

# Create-or-update a Vendor Part rel row carrying the vendor SKU and pricing.
# Mirrors src/aras_mcp/importers/datasheets.py:_upsert_vendor_part.
# Idempotent on (source_id=vendor, related_id=mp).
function Upsert-VendorPart {
    param(
        [Parameter(Mandatory)] [string]$Vendor,        # 'DigiKey' / 'Mouser'
        [Parameter(Mandatory)] [string]$MpId,
        [Parameter(Mandatory)] [string]$CatalogNumber, # vendor SKU
        $UnitPrice     = $null,
        [string]$Currency       = $null,
        $MinOrderQty   = $null,
        $Availability  = $null,
        [string]$ProductStatus  = $null
    )
    if ($DryRun) { return }
    $vendorId = $VendorIds[$Vendor]
    if (-not $vendorId) {
        Write-Log WARN "Upsert-VendorPart: unknown vendor '$Vendor'"
        return
    }
    if ([string]::IsNullOrWhiteSpace($CatalogNumber)) {
        return  # No SKU to anchor on
    }
    $catTrim = $CatalogNumber.Substring(0, [Math]::Min($CatalogNumber.Length, 32))
    $nowIso  = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    $existing = Invoke-RestMethod `
        -Uri "$odataBase/Vendor%20Part?`$filter=source_id eq '$vendorId' and related_id eq '$MpId'&`$select=id,catalog_number,unit_price,currency,min_order_qty,availability,product_status&`$top=1" `
        -Headers $arasHeaders -TimeoutSec 30

    if ($existing.value -and $existing.value.Count -gt 0) {
        $row = $existing.value[0]
        $relId = $row.id
        $patch = @{}
        if ($row.catalog_number -ne $catTrim) { $patch.catalog_number = $catTrim }
        if ($null -ne $UnitPrice) {
            $patch.unit_price = $UnitPrice
            $patch.pricing_updated_on = $nowIso
        }
        if ($Currency -and $Currency -ne $row.currency) {
            $patch.currency = $Currency.Substring(0, [Math]::Min($Currency.Length, 8))
        }
        if ($null -ne $MinOrderQty -and $MinOrderQty -ne $row.min_order_qty) {
            $patch.min_order_qty = $MinOrderQty
        }
        if ($null -ne $Availability -and $Availability -ne $row.availability) {
            $patch.availability = $Availability
        }
        if (-not [string]::IsNullOrWhiteSpace($ProductStatus) -and $ProductStatus -ne $row.product_status) {
            $patch.product_status = $ProductStatus.Substring(0, [Math]::Min($ProductStatus.Length, 32))
        }
        if ($patch.Count -gt 0) {
            $body = $patch | ConvertTo-Json -Compress
            Invoke-RestMethod -Method Patch -Uri "$odataBase/Vendor%20Part('$relId')" `
                -Headers ($arasHeaders + @{ 'Content-Type' = 'application/json' }) `
                -Body $body -TimeoutSec 30 | Out-Null
        }
        return $relId
    }

    $body = @{
        source_id      = $vendorId
        related_id     = $MpId
        catalog_number = $catTrim
    }
    if ($null -ne $UnitPrice) {
        $body.unit_price = $UnitPrice
        $body.pricing_updated_on = $nowIso
    }
    if ($Currency) {
        $body.currency = $Currency.Substring(0, [Math]::Min($Currency.Length, 8))
    }
    if ($null -ne $MinOrderQty) { $body.min_order_qty = $MinOrderQty }
    if ($null -ne $Availability) { $body.availability = $Availability }
    if (-not [string]::IsNullOrWhiteSpace($ProductStatus)) {
        $body.product_status = $ProductStatus.Substring(0, [Math]::Min($ProductStatus.Length, 32))
    }

    $json = $body | ConvertTo-Json -Compress
    $resp = Invoke-RestMethod -Method Post -Uri "$odataBase/Vendor%20Part" `
        -Headers ($arasHeaders + @{ 'Content-Type' = 'application/json' }) `
        -Body $json -TimeoutSec 30
    return $resp.id
}

# ----------------------------------------------------------------------------
# Main loop
# ----------------------------------------------------------------------------

$queued = Invoke-RestMethod `
    -Uri "$odataBase/DatasheetFetchRequest?`$filter=request_state eq 'Queued'&`$select=id,manufacturer_part_id,vendor,attempt_count,file_id&`$top=$Limit&`$orderby=created_on" `
    -Headers $arasHeaders -TimeoutSec 30
Write-Log INFO "Queued rows to drain: $($queued.value.Count)"

$ok = 0; $failed = 0; $skipped = 0
$nowIso = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

foreach ($req in $queued.value) {
    $rowId = $req.id
    $mpId = $req.'manufacturer_part_id@aras.id'
    $mfrPn = $req.'manufacturer_part_id@aras.keyed_name'
    $attempts = [int]$req.attempt_count
    $existingFileId = $req.file_id  # set on requests that already attached a PDF in a prior run

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

        # Always query BOTH vendors so each gets its own Vendor Part row
        # with its own SKU + pricing. Mouser is not a "fallback" — it's a
        # parallel source. The downstream MP.unit_price rollup takes
        # min() across both, so the cheaper vendor wins the cost field.
        $dkLookup = Get-DigiKeyLookup -MfrPn $mfrPn
        $msLookup = $null
        if ($msApiKey) {
            $msLookup = Get-MouserLookup -MfrPn $mfrPn
        }

        $datasheetUrl = $null
        $datasheetSrc = $null   # 'DigiKey' or 'Mouser'
        if ($dkLookup -and $dkLookup.url) {
            $datasheetUrl = $dkLookup.url
            $datasheetSrc = 'DigiKey'
        } elseif ($msLookup -and $msLookup.url) {
            $datasheetUrl = $msLookup.url
            $datasheetSrc = 'Mouser'
        }

        # Upsert Vendor Part rows for every vendor that gave us a SKU
        # (with or without pricing). Pricing-only writes are useful for the
        # cost rollup; SKU-only writes give engineers a stable catalog link.
        $vendorWritten = 0
        if ($dkLookup -and $dkLookup.dk_pn) {
            Upsert-VendorPart -Vendor 'DigiKey' -MpId $mpId `
                -CatalogNumber $dkLookup.dk_pn `
                -UnitPrice $dkLookup.unit_price `
                -Currency  $dkLookup.currency `
                -MinOrderQty $dkLookup.min_order_qty | Out-Null
            $vendorWritten++
        }
        if ($msLookup -and $msLookup.mouser_pn) {
            Upsert-VendorPart -Vendor 'Mouser' -MpId $mpId `
                -CatalogNumber $msLookup.mouser_pn `
                -UnitPrice     $msLookup.unit_price `
                -Currency      $msLookup.currency `
                -MinOrderQty   $msLookup.min_order_qty `
                -Availability  $msLookup.availability `
                -ProductStatus $msLookup.product_status | Out-Null
            $vendorWritten++
        }

        # Datasheet attach. Skip if this request already has file_id set
        # (re-queued for price-only retry) OR if no datasheet URL came back.
        $fileId = $existingFileId
        if (-not $datasheetUrl) {
            if ($vendorWritten -eq 0) {
                Update-Request -RowId $rowId -Body @{
                    request_state = 'Failed'
                    last_error    = 'no datasheet URL or SKU from DigiKey/Mouser'
                }
                Write-Log SKIP "$mfrPn - no vendor match"
                $skipped++; continue
            }
            # Pricing/SKU only — still a useful Completed.
            Update-Request -RowId $rowId -Body @{
                request_state = 'Completed'
                completed_on  = $nowIso
                last_error    = $null
            }
            Write-Log OK "$mfrPn - vendor row(s) created (no datasheet)"
            $ok++; continue
        }

        # Avoid downloading + attaching a duplicate datasheet. Aras tracks
        # the MP↔File attachment, so check that first; only fetch if the
        # MP truly has no file yet AND this request row doesn't already
        # carry a file_id from a prior partial run.
        if (-not $fileId -and (Test-MpHasDatasheet -MpId $mpId)) {
            Write-Log INFO "$mfrPn - already has a datasheet attached; skipping download"
        } elseif (-not $fileId) {
            $bytes = Download-Pdf -Url $datasheetUrl
            $filename = Convert-MpnToSafeFilename -Mpn $mfrPn
            $fileId = Upload-FileToVault -Bytes $bytes -Filename $filename
            New-ManufacturerPartFile -MpId $mpId -FileId $fileId | Out-Null
        } else {
            Write-Log INFO "$mfrPn - reusing existing file_id $existingFileId"
        }
        Update-Request -RowId $rowId -Body @{
            request_state = 'Completed'
            file_id       = $fileId
            datasheet_url = $datasheetUrl
            completed_on  = $nowIso
            last_error    = $null
        }
        if ($null -ne $dkLookup.unit_price -or $null -ne $msLookup.unit_price) {
            Write-Log OK "$mfrPn - datasheet + pricing ($datasheetSrc)"
        } else {
            Write-Log OK "$mfrPn - datasheet only ($datasheetSrc, no pricing)"
        }
        $ok++
    } catch {
        $exType = $_.Exception.GetType().Name
        $exMsg = $_.Exception.Message
        # Try to get the inner response body for HTTP errors
        $detail = ''
        try {
            if ($_.Exception.Response) {
                $sr = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())
                $detail = $sr.ReadToEnd()
                $sr.Close()
            }
        } catch {}
        $msg = "$exType`: $exMsg"
        if ($detail) { $msg += " body=$($detail.Substring(0, [Math]::Min($detail.Length, 400)))" }
        Update-Request -RowId $rowId -Body @{ request_state = 'Failed'; last_error = $msg }
        Write-Log ERROR "$mfrPn - $msg"
        $failed++
    }
}

Write-Log INFO "Done: ok=$ok failed=$failed skipped=$skipped of $($queued.value.Count)"
exit ([int]($failed -gt 0))
