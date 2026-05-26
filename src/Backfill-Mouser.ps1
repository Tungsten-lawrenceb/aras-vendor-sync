<#
.SYNOPSIS
    Walk every current Manufacturer Part and create/update its Mouser
    Vendor Part row with catalog_number + pricing.

.DESCRIPTION
    Standalone backfill, NOT queue-driven. For each MP:
      1. Look up the MPN on Mouser via the Search API.
      2. If a match is found, Upsert-VendorPart with Mouser as the vendor.
      3. Mouser pricing lands in its OWN Vendor Part row (source_id =
         Mouser Electronics), distinct from any DigiKey row. The
         MP.unit_price rollup (MCP_RecomputeMPUnitPrice) takes min()
         across both vendors, so the cheaper one wins.

    Idempotent — re-running this is safe. Existing Mouser rows get
    patched if catalog_number / unit_price changed.

    Reuses Refresh-VendorPricing.ps1's config + OAuth pattern.

.PARAMETER ConfigPath
    Aras config (defaults to C:\ProgramData\AarasVendorSync\config.json).

.PARAMETER Limit
    Max MPs to process. 0 = no limit.

.PARAMETER MpFilter
    Optional OData $filter clause appended to the MP query, e.g.
    "item_number eq 'BMP390'". Lets you narrow a test run.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$env:ProgramData\AarasVendorSync\config.json",
    [int]$Limit = 0,
    [string]$MpFilter = ''
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Mouser allows 30 requests/min on Search; throttle to ~28/min.
. "$PSScriptRoot\Vendor-Throttle.ps1"

# ----------------------------------------------------------------------------
# Config + logging
# ----------------------------------------------------------------------------
if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath" }
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$logDir = $config.log.directory
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logPath = Join-Path $logDir ("{0:yyyy-MM-dd}-mouser-backfill.log" -f (Get-Date))

function Write-Log {
    param([ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'SKIP')] [string]$Level, [string]$Message)
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1,-5} {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $logPath -Value $line
    Write-Host $line
}

Write-Log INFO "=== Backfill-Mouser starting (Limit=$Limit MpFilter='$MpFilter') ==="

# ----------------------------------------------------------------------------
# Aras OAuth + endpoints
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
$odataBase = "$($config.aras.url.TrimEnd('/'))/server/odata"
Write-Log OK "Aras token acquired"

# ----------------------------------------------------------------------------
# Mouser config
# ----------------------------------------------------------------------------
$apiCfg = Invoke-RestMethod -Uri "$odataBase/MPN_API_Config?`$select=vendor,client_id,client_secret,default_endpoint" `
    -Headers $arasHeaders -TimeoutSec 30
$msCfg = $apiCfg.value | Where-Object { $_.vendor -eq 'Mouser' } | Select-Object -First 1
if (-not $msCfg -or -not $msCfg.client_secret) {
    Write-Log ERROR "No Mouser API key in MPN_API_Config"; exit 1
}
$msApiKey = $msCfg.client_secret
$msEndpoint = ($msCfg.default_endpoint -as [string])
if (-not $msEndpoint) { $msEndpoint = 'https://api.mouser.com' }
$msEndpoint = $msEndpoint.TrimEnd('/')
Write-Log OK "Mouser API key loaded"

$VendorIds = @{
    DigiKey = 'E94E47CFF55149C9B16BBD2939D861C9'
    Mouser  = '947C86F817C2454D9C3AE7062D3E0669'
}

# ----------------------------------------------------------------------------
# Mouser lookup + Vendor Part upsert (mirrors Drain-DatasheetQueue.ps1)
# ----------------------------------------------------------------------------
function Get-MpnVariants {
    param([string]$Mpn, [int]$Max = 6)
    $out = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $candidates = @()
    $candidates += $Mpn
    if ($Mpn -match '^(.+),([A-Za-z0-9]+)$') { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)-T\d?$')           { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)-TR$')             { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)#TRPBF$')          { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)#PBF$')            { $candidates += $Matches[1] }
    if ($Mpn -match '^(.+)TR$')              { $candidates += $Matches[1] }
    $collapsed = ($Mpn -replace '\s+', '')
    if ($collapsed -ne $Mpn)                 { $candidates += $collapsed }
    if ($Mpn -match '^(.+)\s+\S+$')          { $candidates += $Matches[1] }
    foreach ($c in $candidates) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        $t = $c.Trim()
        if ($seen.Add($t)) { $out.Add($t) }
        if ($out.Count -ge $Max) { break }
    }
    return $out.ToArray()
}

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

function Get-MouserLookup {
    param([string]$MfrPn)
    $variants = Get-MpnVariants -Mpn $MfrPn
    foreach ($v in $variants) {
        $body = @{
            SearchByPartRequest = @{
                mouserPartNumber  = $v
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
        } catch { continue }
        if ($resp.Errors -and $resp.Errors.Count -gt 0) { continue }
        $parts = $resp.SearchResults.Parts
        if (-not $parts) { continue }

        $targetNorm = ($v -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
        # STRICT: only accept exact alnum-normalized match against the
        # variant we queried. Mouser returns substring matches for
        # non-stocked MPNs and the previous "startswith"/"first hit"
        # fallbacks created bogus Vendor Part rows.
        $exact = @($parts | Where-Object {
                ($_.ManufacturerPartNumber -replace '[^A-Za-z0-9]', '').ToUpperInvariant() -eq $targetNorm
            })
        $chosen = $null
        foreach ($p in $exact) { if ($p.DataSheetUrl) { $chosen = $p; break } }
        if (-not $chosen -and $exact.Count -gt 0) { $chosen = $exact[0] }
        if (-not $chosen) { continue }

        $unitPrice = $null; $currency = $null
        if ($chosen.PriceBreaks) {
            foreach ($pb in $chosen.PriceBreaks) {
                $up = _MouserParseMoney $pb.Price
                if ($null -ne $up) { $unitPrice = $up; $currency = $pb.Currency; break }
            }
        }
        $moq = 0
        if (-not [int]::TryParse(($chosen.Min -as [string]), [ref]$moq)) { $moq = $null }

        $availability  = _MouserParseAvailability $chosen.Availability
        $productStatus = $chosen.LifecycleStatus
        $out = @{
            mouser_pn   = $chosen.MouserPartNumber
            matched_mpn = $chosen.ManufacturerPartNumber
            url         = $chosen.DataSheetUrl
        }
        if ($null -ne $unitPrice)    { $out.unit_price = $unitPrice; $out.currency = $currency }
        if ($moq)                    { $out.min_order_qty = $moq }
        if ($null -ne $availability) { $out.availability = $availability }
        if (-not [string]::IsNullOrWhiteSpace($productStatus)) {
            $out.product_status = $productStatus
        }
        return $out
    }
    return $null
}

function Upsert-VendorPart {
    param(
        [Parameter(Mandatory)] [string]$Vendor,
        [Parameter(Mandatory)] [string]$MpId,
        [Parameter(Mandatory)] [string]$CatalogNumber,
        $UnitPrice = $null, [string]$Currency = $null, $MinOrderQty = $null,
        $Availability = $null, [string]$ProductStatus = $null
    )
    $vendorId = $VendorIds[$Vendor]
    if (-not $vendorId -or [string]::IsNullOrWhiteSpace($CatalogNumber)) { return $null }
    $catTrim = $CatalogNumber.Substring(0, [Math]::Min($CatalogNumber.Length, 32))
    $nowIso  = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    $existing = Invoke-RestMethod `
        -Uri "$odataBase/Vendor%20Part?`$filter=source_id eq '$vendorId' and related_id eq '$MpId'&`$select=id,catalog_number,unit_price,currency,min_order_qty,availability,product_status&`$top=1" `
        -Headers $arasHeaders -TimeoutSec 30

    if ($existing.value -and $existing.value.Count -gt 0) {
        $row = $existing.value[0]
        $patch = @{}
        if ($row.catalog_number -ne $catTrim) { $patch.catalog_number = $catTrim }
        if ($null -ne $UnitPrice) {
            $patch.unit_price = $UnitPrice; $patch.pricing_updated_on = $nowIso
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
            Invoke-RestMethod -Method Patch -Uri "$odataBase/Vendor%20Part('$($row.id)')" `
                -Headers ($arasHeaders + @{ 'Content-Type' = 'application/json' }) `
                -Body $body -TimeoutSec 30 | Out-Null
            return @{ id = $row.id; action = 'patched' }
        }
        return @{ id = $row.id; action = 'unchanged' }
    }

    $body = @{ source_id = $vendorId; related_id = $MpId; catalog_number = $catTrim }
    if ($null -ne $UnitPrice) { $body.unit_price = $UnitPrice; $body.pricing_updated_on = $nowIso }
    if ($Currency)            { $body.currency = $Currency.Substring(0, [Math]::Min($Currency.Length, 8)) }
    if ($null -ne $MinOrderQty)  { $body.min_order_qty = $MinOrderQty }
    if ($null -ne $Availability) { $body.availability = $Availability }
    if (-not [string]::IsNullOrWhiteSpace($ProductStatus)) {
        $body.product_status = $ProductStatus.Substring(0, [Math]::Min($ProductStatus.Length, 32))
    }
    $json = $body | ConvertTo-Json -Compress
    $resp = Invoke-RestMethod -Method Post -Uri "$odataBase/Vendor%20Part" `
        -Headers ($arasHeaders + @{ 'Content-Type' = 'application/json' }) `
        -Body $json -TimeoutSec 30
    return @{ id = $resp.id; action = 'created' }
}

# ----------------------------------------------------------------------------
# Walk MPs
# ----------------------------------------------------------------------------
$filterClause = "is_current eq '1'"
if ($MpFilter) { $filterClause += " and ($MpFilter)" }
$topClause = if ($Limit -gt 0) { "&`$top=$Limit" } else { '' }
$encoded = [System.Uri]::EscapeDataString($filterClause)
$mps = Invoke-RestMethod -Uri "$odataBase/Manufacturer%20Part?`$filter=$encoded&`$select=id,item_number$topClause" `
    -Headers $arasHeaders -TimeoutSec 60
$total = $mps.value.Count
Write-Log INFO "MPs to scan: $total"

$created = 0; $patched = 0; $unchanged = 0; $miss = 0; $err = 0
foreach ($mp in $mps.value) {
    $mpId = $mp.id
    $mfrPn = $mp.item_number
    if (-not $mfrPn) { continue }
    try {
        $lookup = Get-MouserLookup -MfrPn $mfrPn
        if (-not $lookup -or -not $lookup.mouser_pn) {
            $miss++
            continue
        }
        $res = Upsert-VendorPart -Vendor 'Mouser' -MpId $mpId `
            -CatalogNumber $lookup.mouser_pn `
            -UnitPrice     $lookup.unit_price `
            -Currency      $lookup.currency `
            -MinOrderQty   $lookup.min_order_qty `
            -Availability  $lookup.availability `
            -ProductStatus $lookup.product_status
        if ($res.action -eq 'created')    { $created++   }
        elseif ($res.action -eq 'patched')  { $patched++   }
        else                                { $unchanged++ }
        $price = if ($null -ne $lookup.unit_price) { "`$$($lookup.unit_price)" } else { 'no-price' }
        Write-Log OK "$mfrPn => $($lookup.mouser_pn) ($price) [$($res.action)]"
    } catch {
        $err++
        Write-Log ERROR "$mfrPn - $($_.Exception.GetType().Name): $($_.Exception.Message)"
    }
}

Write-Log INFO "Done: created=$created patched=$patched unchanged=$unchanged miss=$miss errors=$err of $total"
exit ([int]($err -gt 0))
