<#
.SYNOPSIS
    Shared per-vendor rate-limit throttle. Dot-source from any script
    that calls the Digi-Key or Mouser APIs.

.DESCRIPTION
    Published limits:
      - Mouser Search API: 30 requests/min (their docs). Free tier
        accounts hit this fast. Default interval here is 2.1 sec ->
        ~28 calls/min, leaving headroom.
      - Digi-Key V4 Production API: typical app limit is 1000-10000
        requests/day depending on subscription, and rough burst limit
        of ~10 req/sec. We default to 600 ms (100 req/min) to be safe.

    Both APIs return HTTP 429 when limits are hit. Invoke-VendorApi
    honors the Retry-After header (or sleeps 60 sec if absent) and
    retries once. Persistent 429 is treated as an error.

.USAGE
    . "$PSScriptRoot\Vendor-Throttle.ps1"

    # Either of these usage patterns works:
    Wait-ApiInterval -Api 'Mouser'
    $r = Invoke-RestMethod ...

    # OR (preferred -- handles 429 too):
    $r = Invoke-VendorApi -Api 'DigiKey' -Method Get -Uri $url -Headers $h -TimeoutSec 30
#>

# Script-scoped state -- one set per script invocation.
$script:LastApiCall     = @{ DigiKey = [DateTime]::MinValue; Mouser = [DateTime]::MinValue }
# Conservative defaults under each vendor's published limits:
#   DigiKey: documented 240/min burst on most tiers; we sustain 40/min
#     (1500ms) because our app's actual cap is unknown and we've seen
#     429s at 60/min.
#   Mouser: documented 30/min cap on the Search API.
$script:ApiMinIntervalMs = @{ DigiKey = 1500; Mouser = 2100 }

# Allow overrides via env vars for tuning without code changes.
if ($env:DIGIKEY_API_INTERVAL_MS) { $script:ApiMinIntervalMs.DigiKey = [int]$env:DIGIKEY_API_INTERVAL_MS }
if ($env:MOUSER_API_INTERVAL_MS)  { $script:ApiMinIntervalMs.Mouser  = [int]$env:MOUSER_API_INTERVAL_MS }

function Wait-ApiInterval {
    param([Parameter(Mandatory)] [ValidateSet('DigiKey','Mouser')] [string]$Api)
    $minMs = $script:ApiMinIntervalMs[$Api]
    if (-not $minMs) { $script:LastApiCall[$Api] = Get-Date; return }
    $elapsed = ((Get-Date) - $script:LastApiCall[$Api]).TotalMilliseconds
    if ($elapsed -lt $minMs) {
        Start-Sleep -Milliseconds ([int]($minMs - $elapsed))
    }
    $script:LastApiCall[$Api] = Get-Date
}

function Invoke-VendorApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ValidateSet('DigiKey','Mouser')] [string]$Api,
        [Parameter(Mandatory)] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [hashtable]$Headers,
        [string]$ContentType,
        $Body,
        [int]$TimeoutSec = 30,
        [int]$MaxRetries = 1   # one 429 retry by default
    )
    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        Wait-ApiInterval -Api $Api
        try {
            $params = @{
                Method     = $Method
                Uri        = $Uri
                TimeoutSec = $TimeoutSec
            }
            if ($Headers)         { $params.Headers     = $Headers }
            if ($ContentType)     { $params.ContentType = $ContentType }
            if ($null -ne $Body)  { $params.Body        = $Body }
            return Invoke-RestMethod @params
        } catch {
            $resp = $_.Exception.Response
            $status = $null
            if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode }
            if ($status -eq 429 -and $attempt -lt $MaxRetries) {
                # Honor Retry-After if present (seconds), but CAP at 60 sec.
                # DigiKey sometimes returns Retry-After in the thousands
                # (their daily quota reset window) and we'd otherwise stall
                # the whole script for hours. If the server really wants
                # us to wait longer, propagating the 429 is better — the
                # caller (e.g. verify) will mark the row "uncertain" and
                # the user can re-run later.
                $wait = 60
                if ($resp -and $resp.Headers) {
                    $ra = $null
                    try { $ra = $resp.Headers['Retry-After'] } catch {}
                    if ($ra) {
                        $n = 0
                        if ([int]::TryParse($ra, [ref]$n)) { $wait = [Math]::Max(1, [Math]::Min($n, 60)) }
                    }
                }
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}
