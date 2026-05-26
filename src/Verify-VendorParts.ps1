<#
.SYNOPSIS
    Walk every existing Vendor Part row and verify the vendor still
    claims our MP's MPN. Delete rows that fail the strict alnum-equal
    test (i.e. were created by an earlier loose-fallback bug).

.DESCRIPTION
    For each current Vendor Part row:
      1. Identify the vendor (DigiKey / Mouser via source_id).
      2. Call the vendor API for the MP's item_number (with the same
         loose-MPN variant generation the drainer uses).
      3. If ANY variant returns a hit whose ManufacturerPartNumber
         alnum-equals our variant, the row is legit -- keep.
      4. Otherwise DELETE the row. The Vendor Part rel referenced
         pricing/SKU that doesn't actually belong to our MP.

    Idempotent and conservative: any API hiccup (timeout, 5xx) is
    treated as "uncertain" -- the row is kept. Only definitive
    "vendor responded but no MPN matches" triggers deletion.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$env:ProgramData\AarasVendorSync\config.json",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Per-vendor throttle so a full scan doesn't trip Mouser's 30/min limit.
. "$PSScriptRoot\Vendor-Throttle.ps1"

if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath" }
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$logDir = $config.log.directory
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logPath = Join-Path $logDir ("{0:yyyy-MM-dd}-verify.log" -f (Get-Date))

function Write-Log {
    param([ValidateSet('INFO','WARN','ERROR','OK','KEEP','DELETE')] [string]$Level, [string]$Message)
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1,-6} {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $logPath -Value $line
    Write-Host $line
}

Write-Log INFO "=== Verify-VendorParts starting (DryRun=$DryRun) ==="

# ---- Aras OAuth ----
function Get-ArasToken {
    param($A)
    $r = Invoke-RestMethod -Method Post -Uri "$($A.url.TrimEnd('/'))/OAuthServer/connect/token" -Body @{
        grant_type='password'; scope='Innovator'; client_id='IOMApp'
        username=$A.username; password=$A.password_md5; database=$A.database
    } -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 30
    return $r.access_token
}
$script:arasTokenAcquiredAt = Get-Date
$arasToken = Get-ArasToken $config.aras
$arasHeaders = @{ Authorization = "Bearer $arasToken"; Accept = 'application/json' }
$odataBase = "$($config.aras.url.TrimEnd('/'))/server/odata"
Write-Log OK "Aras token acquired"

# Long verify runs can outlast the Aras OAuth token (~1 hr default).
# Refresh every 30 min defensively.
function Refresh-ArasTokenIfStale {
    $age = (Get-Date) - $script:arasTokenAcquiredAt
    if ($age.TotalMinutes -lt 30) { return }
    $script:arasToken = Get-ArasToken $config.aras
    $script:arasHeaders = @{ Authorization = "Bearer $script:arasToken"; Accept = 'application/json' }
    $script:arasTokenAcquiredAt = Get-Date
    Write-Log INFO "Aras token refreshed (age was $([int]$age.TotalMinutes) min)"
}

# ---- Vendor configs ----
$apiCfg = Invoke-RestMethod -Uri "$odataBase/MPN_API_Config?`$select=vendor,client_id,client_secret,default_endpoint" -Headers $arasHeaders -TimeoutSec 30
$dkCfg = $apiCfg.value | Where-Object { $_.vendor -eq 'DigiKey' } | Select-Object -First 1
$msCfg = $apiCfg.value | Where-Object { $_.vendor -eq 'Mouser'  } | Select-Object -First 1

$dkEndpoint = $null; $dkHeaders = $null
if ($dkCfg -and $dkCfg.client_id -and $dkCfg.client_secret) {
    $dkEndpoint = ($dkCfg.default_endpoint -as [string]); if (-not $dkEndpoint) { $dkEndpoint = 'https://api.digikey.com' }
    $dkEndpoint = $dkEndpoint.TrimEnd('/')
    $tok = Invoke-RestMethod -Method Post -Uri "$dkEndpoint/v1/oauth2/token" -ContentType 'application/x-www-form-urlencoded' -Body @{
        grant_type='client_credentials'; client_id=$dkCfg.client_id; client_secret=$dkCfg.client_secret
    } -TimeoutSec 30
    $dkHeaders = @{ Authorization = "Bearer $($tok.access_token)"; 'X-DIGIKEY-Client-Id' = $dkCfg.client_id; Accept = 'application/json' }
    Write-Log OK "Digi-Key token acquired"
}
$msEndpoint = $null; $msApiKey = $null
if ($msCfg -and $msCfg.client_secret) {
    $msEndpoint = ($msCfg.default_endpoint -as [string]); if (-not $msEndpoint) { $msEndpoint = 'https://api.mouser.com' }
    $msEndpoint = $msEndpoint.TrimEnd('/')
    $msApiKey = $msCfg.client_secret
    Write-Log OK "Mouser key loaded"
}

$VendorIds = @{
    'E94E47CFF55149C9B16BBD2939D861C9' = 'DigiKey'
    '947C86F817C2454D9C3AE7062D3E0669' = 'Mouser'
}

function Get-MpnVariants {
    param([string]$Mpn, [int]$Max = 6)
    $out = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $candidates = @($Mpn)
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

function _Norm { param([string]$S) return (($S -replace '[^A-Za-z0-9]', '').ToUpperInvariant()) }

# Returns one of: 'match', 'no-match', 'uncertain'
function Test-DigiKeyHasMpn {
    param([string]$Mpn)
    if (-not $dkHeaders) { return 'uncertain' }
    $variants = Get-MpnVariants -Mpn $Mpn
    foreach ($v in $variants) {
        $vNorm = _Norm $v
        $encoded = [System.Uri]::EscapeDataString($v)
        # ProductDetails: a 404 is a clean "no exact MPN" signal, not an
        # error. Silently move on to keyword search.
        try {
            $r = Invoke-VendorApi -Api 'DigiKey' -Method Get `
                -Uri "$dkEndpoint/products/v4/search/$encoded/productdetails" `
                -Headers $dkHeaders -TimeoutSec 30
            if ($r.Product -and (_Norm $r.Product.ManufacturerProductNumber) -eq $vNorm) { return 'match' }
        } catch {}
        # Keyword search: also tolerate 404. Only treat a persistent
        # non-2xx as "uncertain" -- meaning we don't know whether the
        # vendor has this MPN, so keep the row.
        try {
            $kwBody = @{ Keywords = $v; Limit = 10; Offset = 0 } | ConvertTo-Json -Compress
            $kw = Invoke-VendorApi -Api 'DigiKey' -Method Post `
                -Uri "$dkEndpoint/products/v4/search/keyword" `
                -Headers ($dkHeaders + @{ 'Content-Type' = 'application/json' }) `
                -ContentType 'application/json' `
                -Body $kwBody -TimeoutSec 30
            if ($kw.Products) {
                foreach ($p in $kw.Products) {
                    if ((_Norm $p.ManufacturerProductNumber) -eq $vNorm) { return 'match' }
                }
            }
        } catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 404) { continue }   # legit "no result" - next variant
            return 'uncertain'
        }
    }
    return 'no-match'
}

function Test-MouserHasMpn {
    param([string]$Mpn)
    if (-not $msApiKey) { return 'uncertain' }
    $variants = Get-MpnVariants -Mpn $Mpn
    foreach ($v in $variants) {
        $vNorm = _Norm $v
        $body = @{ SearchByPartRequest = @{ mouserPartNumber=$v; partSearchOptions='string' } } | ConvertTo-Json -Compress -Depth 4
        try {
            $r = Invoke-VendorApi -Api 'Mouser' -Method Post `
                -Uri "$msEndpoint/api/v1/search/partnumber?apiKey=$msApiKey" `
                -Headers @{ Accept='application/json' } -ContentType 'application/json' -Body $body -TimeoutSec 30
        } catch { return 'uncertain' }
        if ($r.Errors -and $r.Errors.Count -gt 0) { continue }
        if (-not $r.SearchResults -or -not $r.SearchResults.Parts) { continue }
        foreach ($p in $r.SearchResults.Parts) {
            if ((_Norm $p.ManufacturerPartNumber) -eq $vNorm) { return 'match' }
        }
    }
    return 'no-match'
}

# ---- Walk every Vendor Part row ----
$vps = Invoke-RestMethod -Uri "$odataBase/Vendor%20Part?`$select=id,source_id,related_id,catalog_number&`$top=10000" -Headers $arasHeaders -TimeoutSec 60
Write-Log INFO "Vendor Part rows: $($vps.value.Count)"

$kept = 0; $deleted = 0; $uncertain = 0; $skipped = 0
foreach ($vp in $vps.value) {
    $sId = $vp.'source_id@aras.id'; if (-not $sId) { $sId = $vp.source_id }
    $rId = $vp.'related_id@aras.id'; if (-not $rId) { $rId = $vp.related_id }
    $vendor = $VendorIds[$sId]
    if (-not $vendor) { $skipped++; continue }

    Refresh-ArasTokenIfStale
    # Pull MP item_number
    try {
        $mp = Invoke-RestMethod -Uri "$odataBase/Manufacturer%20Part('$rId')?`$select=item_number" -Headers $script:arasHeaders -TimeoutSec 30
    } catch { Write-Log WARN "MP fetch failed for $rId"; $skipped++; continue }
    $mfrPn = $mp.item_number
    if (-not $mfrPn) { $skipped++; continue }

    $verdict = switch ($vendor) {
        'DigiKey' { Test-DigiKeyHasMpn -Mpn $mfrPn }
        'Mouser'  { Test-MouserHasMpn  -Mpn $mfrPn }
    }
    if ($verdict -eq 'match') {
        $kept++
    }
    elseif ($verdict -eq 'no-match') {
        Write-Log DELETE "$mfrPn [$vendor] - catalog=$($vp.catalog_number) - no MPN match, deleting Vendor Part $($vp.id)"
        $deleted++
        if (-not $DryRun) {
            try {
                Invoke-RestMethod -Method Delete -Uri "$odataBase/Vendor%20Part('$($vp.id)')" -Headers $arasHeaders -TimeoutSec 30 | Out-Null
            } catch { Write-Log ERROR "delete failed for $($vp.id): $($_.Exception.Message)" }
        }
    }
    else {
        Write-Log WARN "$mfrPn [$vendor] - uncertain (API hiccup); keeping"
        $uncertain++
    }
}

Write-Log INFO "Done: kept=$kept deleted=$deleted uncertain=$uncertain skipped=$skipped"
exit ([int]($deleted -gt 0 -and $DryRun))
