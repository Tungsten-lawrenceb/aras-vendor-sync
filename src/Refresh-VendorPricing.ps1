<#
.SYNOPSIS
    Refresh per-vendor pricing on every Vendor Part row in Aras Innovator.

.DESCRIPTION
    Reads vendor credentials from MPN_API_Config rows in Aras, walks
    every Vendor Part with a catalog_number set, calls the configured
    vendor's API to fetch fresh pricing, and PATCHes the row. The
    server-side MCP_RecomputeMPUnitPrice Method (bound to Vendor
    Part.onAfterUpdate) automatically rolls each Vendor Part's new
    unit_price into the parent Manufacturer Part's unit_price as
    min across all vendors.

    Designed to run as a Windows Scheduled Task on the Aras VM. See
    scripts/Install-ScheduledTask.ps1.

.PARAMETER ConfigPath
    Path to the JSON config file. Defaults to
    C:\ProgramData\AarasVendorSync\config.json.

.PARAMETER DryRun
    Report what would be patched without writing anything to Aras.

.EXAMPLE
    .\Refresh-VendorPricing.ps1 -DryRun

.EXAMPLE
    .\Refresh-VendorPricing.ps1 -ConfigPath D:\custom-config.json

.NOTES
    Author: Tungsten Collaborative
    Repo:   https://github.com/Tungsten-lawrenceb/aras-vendor-sync
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$env:ProgramData\AarasVendorSync\config.json",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Per-vendor throttle so the weekly walk doesn't hammer DK above its rate.
. "$PSScriptRoot\Vendor-Throttle.ps1"

# ----------------------------------------------------------------------------
# Config + logging
# ----------------------------------------------------------------------------

if (-not (Test-Path $ConfigPath)) {
    throw "Config not found: $ConfigPath. Copy from config/config.example.json and fill in."
}
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

$logDir = $config.log.directory
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logPath = Join-Path $logDir ("{0:yyyy-MM-dd}.log" -f (Get-Date))

# Ensure the Event Log source exists. New-EventLog requires Admin rights;
# silently skip if we lack them (the file log is the primary surface).
try {
    if (-not [System.Diagnostics.EventLog]::SourceExists($config.log.event_log_source)) {
        New-EventLog -LogName Application -Source $config.log.event_log_source -ErrorAction Stop
    }
} catch {
    # Likely no admin rights, or already exists in a way we can't detect. Move on.
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [ValidateSet('INFO', 'WARN', 'ERROR', 'OK', 'SKIP')] [string]$Level,
        [Parameter(Mandatory)] [string]$Message
    )
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1,-5} {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $logPath -Value $line
    Write-Host $line
}

function Write-EventLogSafe {
    param([string]$Message, [string]$EntryType = 'Information', [int]$EventId = 1000)
    try {
        Write-EventLog -LogName Application -Source $config.log.event_log_source `
            -EntryType $EntryType -EventId $EventId -Message $Message -ErrorAction Stop
    } catch {
        # Falls back to the file log only.
    }
}

Write-Log INFO "=== aras-vendor-sync starting (DryRun=$DryRun) ==="
Write-Log INFO "Config: $ConfigPath"
Write-Log INFO "Aras:   $($config.aras.url) / db=$($config.aras.database) / user=$($config.aras.username)"

# ----------------------------------------------------------------------------
# Aras OAuth
# ----------------------------------------------------------------------------

function Get-ArasToken {
    param([Parameter(Mandatory)] $ArasConfig)
    # Token endpoint is fixed on the standard IIS layout. Discovery via
    # /Server/OAuthServerDiscovery.aspx is supported but adds two extra
    # round-trips; on Tungsten's install the path below is canonical.
    $url = "$($ArasConfig.url.TrimEnd('/'))/OAuthServer/connect/token"
    $body = @{
        grant_type = 'password'
        scope      = 'Innovator'
        client_id  = 'IOMApp'
        username   = $ArasConfig.username
        password   = $ArasConfig.password_md5
        database   = $ArasConfig.database
    }
    try {
        $resp = Invoke-RestMethod -Method Post -Uri $url -Body $body `
            -ContentType 'application/x-www-form-urlencoded' `
            -TimeoutSec 30
    } catch {
        throw "Aras token endpoint failed: $($_.Exception.Message)"
    }
    if (-not $resp.access_token) {
        throw "Aras token response missing access_token"
    }
    return $resp.access_token
}

$arasToken = Get-ArasToken -ArasConfig $config.aras
Write-Log OK "Aras token acquired"

$arasHeaders = @{
    Authorization = "Bearer $arasToken"
    Accept        = 'application/json'
}
$odataBase = "$($config.aras.url.TrimEnd('/'))/server/odata"

# ----------------------------------------------------------------------------
# Read vendor credentials from MPN_API_Config
# ----------------------------------------------------------------------------

function Get-MpnApiConfig {
    $url = "$odataBase/MPN_API_Config?`$select=vendor,client_id,client_secret,default_endpoint"
    $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $arasHeaders -TimeoutSec 30
    $out = @{}
    foreach ($row in $resp.value) {
        $out[$row.vendor] = $row
    }
    return $out
}

$vendorConfigs = Get-MpnApiConfig
Write-Log INFO ("MPN_API_Config rows: " + ($vendorConfigs.Keys -join ', '))

if (-not $vendorConfigs.ContainsKey('DigiKey')) {
    Write-Log ERROR "No DigiKey row in MPN_API_Config. Aborting."
    Write-EventLogSafe -EntryType Error -EventId 2001 `
        -Message "aras-vendor-sync: no DigiKey config in MPN_API_Config"
    exit 1
}
$dkCfg = $vendorConfigs['DigiKey']
$dkEndpoint = if ($dkCfg.default_endpoint) { $dkCfg.default_endpoint } else { 'https://api.digikey.com' }
$dkEndpoint = $dkEndpoint.TrimEnd('/')

# ----------------------------------------------------------------------------
# Digi-Key OAuth (client_credentials)
# ----------------------------------------------------------------------------

function Get-DigiKeyToken {
    param([string]$Endpoint, [string]$ClientId, [string]$ClientSecret)
    $resp = Invoke-RestMethod -Method Post -Uri "$Endpoint/v1/oauth2/token" `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $ClientSecret
        } -TimeoutSec 30
    if (-not $resp.access_token) { throw 'DigiKey token response missing access_token' }
    return $resp.access_token
}

$dkToken = Get-DigiKeyToken -Endpoint $dkEndpoint `
    -ClientId $dkCfg.client_id -ClientSecret $dkCfg.client_secret
Write-Log OK "DigiKey token acquired"

# ----------------------------------------------------------------------------
# Digi-Key ProductDetails (with keyword fallback for "Duplicate Products")
# ----------------------------------------------------------------------------

$dkHeaders = @{
    Authorization              = "Bearer $dkToken"
    'X-DIGIKEY-Client-Id'      = $dkCfg.client_id
    Accept                     = 'application/json'
}

function Get-DigiKeyPricing {
    param([string]$MfrPn)

    # Extract the lowest-MOQ variation's qty-1 price. Falls back to
    # top-level Product.UnitPrice. Mirrors the Python worker's logic.
    function _extract($product) {
        $out = [PSCustomObject]@{
            unit_price          = $null
            currency            = 'USD'
            min_order_qty       = $null
            availability        = $null
            product_status      = $null
            digikey_part_number = $null
            datasheet_url       = $product.DatasheetUrl
        }
        if ($product.ProductStatus -and $product.ProductStatus.Status) {
            $out.product_status = $product.ProductStatus.Status
        }
        if ($product.QuantityAvailable -ne $null) {
            $out.availability = [int]$product.QuantityAvailable
        }
        $chosen = $null
        $chosenMoq = $null
        foreach ($v in $product.ProductVariations) {
            $moqRaw = $v.MinimumOrderQuantity
            $moq = $null
            if ($moqRaw -ne $null -and [int]::TryParse([string]$moqRaw, [ref]([int]$null))) {
                $moq = [int]$moqRaw
            }
            if (-not $chosen -or ($moq -ne $null -and ($chosenMoq -eq $null -or $moq -lt $chosenMoq))) {
                $chosen = $v
                $chosenMoq = $moq
            }
        }
        if ($chosen) {
            $out.digikey_part_number = $chosen.DigiKeyProductNumber
            $out.min_order_qty = $chosenMoq
            $firstBreak = $chosen.StandardPricing | Select-Object -First 1
            if ($firstBreak -and $firstBreak.UnitPrice -ne $null) {
                $out.unit_price = [decimal]$firstBreak.UnitPrice
            }
        }
        if (-not $out.unit_price -and $product.UnitPrice -ne $null) {
            $out.unit_price = [decimal]$product.UnitPrice
        }
        return $out
    }

    # ProductDetails (exact MPN). 404 + "Duplicate Products" needs keyword fallback.
    $encoded = [System.Uri]::EscapeDataString($MfrPn)
    $url = "$dkEndpoint/products/v4/search/$encoded/productdetails"
    Wait-ApiInterval -Api 'DigiKey'
    try {
        $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $dkHeaders -TimeoutSec 30
        if ($resp.Product) {
            return _extract $resp.Product
        }
    } catch {
        # Non-2xx - try keyword search as a fallback (handles
        # "Duplicate Products found for X" 404s).
    }
    Wait-ApiInterval -Api 'DigiKey'
    try {
        $kwBody = @{ Keywords = $MfrPn; Limit = 5; Offset = 0 } | ConvertTo-Json -Compress
        $kw = Invoke-RestMethod -Method Post -Uri "$dkEndpoint/products/v4/search/keyword" `
            -Headers ($dkHeaders + @{ 'Content-Type' = 'application/json' }) `
            -Body $kwBody -TimeoutSec 30
    } catch {
        return $null
    }
    if (-not $kw -or -not $kw.Products) { return $null }
    $targetNorm = ($MfrPn -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    # Prefer exact MPN match on mm.digikey.com CDN (those URLs are reachable
    # from our IP), then any exact match, then any DK CDN hit, then any hit.
    function _isDkCdn($u) { return $u -and $u.ToLower().Contains('mm.digikey.com') }
    $candidates = @($kw.Products)
    $exact = $candidates | Where-Object {
        ($_.ManufacturerProductNumber -replace '[^A-Za-z0-9]', '').ToUpperInvariant() -eq $targetNorm
    }
    $best = $exact | Where-Object { _isDkCdn $_.DatasheetUrl } | Select-Object -First 1
    if (-not $best) { $best = $exact | Where-Object { $_.DatasheetUrl } | Select-Object -First 1 }
    if (-not $best) { $best = $candidates | Where-Object { _isDkCdn $_.DatasheetUrl } | Select-Object -First 1 }
    if (-not $best) { $best = $candidates | Where-Object { $_.DatasheetUrl } | Select-Object -First 1 }
    if (-not $best) { return $null }
    return _extract $best
}

# ----------------------------------------------------------------------------
# List Vendor Part rows to refresh
# ----------------------------------------------------------------------------

function Get-VendorPartRows {
    param([int]$Limit = 500)
    # Only refresh rows attributed to Digi-Key Electronics. Mouser will be
    # added when its refresh path lands.
    $url = "$odataBase/Vendor%20Part?`$filter=catalog_number ne null and source_id eq 'E94E47CFF55149C9B16BBD2939D861C9'" `
        + "&`$select=id,source_id,related_id,catalog_number,unit_price,currency,pricing_updated_on&`$top=$Limit"
    $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $arasHeaders -TimeoutSec 60
    return $resp.value
}

$maxRows = if ($config.limits.max_rows_per_run) { [int]$config.limits.max_rows_per_run } else { 500 }
$rows = Get-VendorPartRows -Limit $maxRows
Write-Log INFO "Vendor Part rows to refresh (Digi-Key): $($rows.Count)"

# ----------------------------------------------------------------------------
# Patch loop
# ----------------------------------------------------------------------------

function Update-VendorPart {
    param([string]$RowId, [hashtable]$Body)
    if ($DryRun) { return }
    $url = "$odataBase/Vendor%20Part('$RowId')"
    $json = $Body | ConvertTo-Json -Depth 4 -Compress
    Invoke-RestMethod -Method Patch -Uri $url `
        -Headers ($arasHeaders + @{ 'Content-Type' = 'application/json' }) `
        -Body $json -TimeoutSec 30 | Out-Null
}

$ok = 0
$failed = 0
$skipped = 0
$nowIso = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

foreach ($row in $rows) {
    $mpId = $row.'related_id@aras.id'
    $mfrPn = $row.'related_id@aras.keyed_name'
    if (-not $mfrPn) {
        Write-Log SKIP "row=$($row.id) - no Mfr PN on related_id"
        $skipped++
        continue
    }
    try {
        $price = Get-DigiKeyPricing -MfrPn $mfrPn
    } catch {
        Write-Log ERROR "$mfrPn - DigiKey error: $($_.Exception.Message)"
        $failed++
        continue
    }
    if (-not $price -or $price.unit_price -eq $null) {
        Write-Log SKIP "$mfrPn - no pricing in DK response"
        $skipped++
        continue
    }

    $patch = @{
        unit_price         = [double]$price.unit_price
        currency           = $price.currency
        pricing_updated_on = $nowIso
    }
    if ($price.min_order_qty -ne $null) { $patch['min_order_qty'] = $price.min_order_qty }
    if ($price.availability  -ne $null) { $patch['availability']  = $price.availability  }
    if ($price.product_status)          { $patch['product_status'] = $price.product_status }
    if ($price.digikey_part_number `
        -and $price.digikey_part_number -ne $row.catalog_number) {
        $patch['catalog_number'] = $price.digikey_part_number
    }

    try {
        Update-VendorPart -RowId $row.id -Body $patch
        $delta = ""
        if ($row.unit_price -ne $null) {
            $delta = " (was $([decimal]$row.unit_price))"
        }
        Write-Log OK ("{0,-28} sku={1,-26} price={2}{3} status={4}" -f `
            $mfrPn, $price.digikey_part_number, $price.unit_price, $delta, $price.product_status)
        $ok++
    } catch {
        Write-Log ERROR "$mfrPn - PATCH failed: $($_.Exception.Message)"
        $failed++
    }
}

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------

$summary = "Done: ok=$ok failed=$failed skipped=$skipped of $($rows.Count)"
Write-Log INFO $summary

$entryType = if ($failed -gt 0) { 'Warning' } else { 'Information' }
$eventId = if ($failed -gt 0) { 2002 } else { 1001 }
Write-EventLogSafe -Message "aras-vendor-sync: $summary" -EntryType $entryType -EventId $eventId

if ($DryRun) {
    Write-Log INFO "DryRun=true - no rows were patched."
}

exit ([int]($failed -gt 0))
