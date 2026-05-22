<#
.SYNOPSIS
    Register the Windows Scheduled Task that runs the aras-vendor-sync
    refresh script weekly on Sundays at 03:00 local time.

.PARAMETER ScriptPath
    Path to Refresh-VendorPricing.ps1. Defaults to the sibling src/ dir.

.PARAMETER ConfigPath
    Path to the JSON config. Defaults to
    C:\ProgramData\AarasVendorSync\config.json.

.PARAMETER TaskName
    Scheduled Task name. Defaults to 'aras-vendor-sync'.

.PARAMETER At
    Time-of-day for the trigger. Defaults to 03:00.

.PARAMETER DayOfWeek
    Day-of-week for the trigger. Defaults to Sunday.

.PARAMETER User
    The user identity the task runs as. Defaults to SYSTEM. Use a
    dedicated service account in production so it shows up in audits.

.EXAMPLE
    # Default: Sundays 03:00 as SYSTEM
    .\Install-ScheduledTask.ps1

.EXAMPLE
    # Daily at 02:00 as the local service account
    .\Install-ScheduledTask.ps1 -DayOfWeek @('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday') -At 02:00
#>

[CmdletBinding()]
param(
    [string]$ScriptPath,
    [string]$ConfigPath = "$env:ProgramData\AarasVendorSync\config.json",
    [string]$TaskName   = 'aras-vendor-sync',
    [string]$At         = '03:00',
    [string[]]$DayOfWeek = @('Sunday'),
    [string]$User       = 'SYSTEM'
)

$ErrorActionPreference = 'Stop'

# Resolve the default ScriptPath inside the body so $PSScriptRoot is
# guaranteed populated. (PS sets $PSScriptRoot when the script is
# invoked via -File, but param-default expressions can evaluate before
# that binding in some invocation contexts.)
if (-not $ScriptPath) {
    $here = $PSScriptRoot
    if (-not $here) { $here = Split-Path $MyInvocation.MyCommand.Path -Parent }
    $ScriptPath = Join-Path (Split-Path $here -Parent) 'src\Refresh-VendorPricing.ps1'
}

if (-not (Test-Path $ScriptPath)) {
    throw "Script not found: $ScriptPath"
}
if (-not (Test-Path $ConfigPath)) {
    Write-Warning "Config not found at $ConfigPath - create it before the next scheduled run."
}

# Build the arguments. PowerShell needs to be invoked with -File and the
# -ConfigPath flag passed through.
# Prefer pwsh (PowerShell 7) if installed; fall back to Windows PowerShell 5.1.
# `??` (null-coalescing) is PS 7+ only - write it the PS 5.1-compatible way.
$psExe = Get-Command pwsh -ErrorAction SilentlyContinue
if (-not $psExe) { $psExe = Get-Command powershell }
$action = New-ScheduledTaskAction `
    -Execute $psExe.Source `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -ConfigPath `"$ConfigPath`""

# Weekly trigger
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $At

# Settings: start when the task is registered, don't pile up if missed runs accumulate,
# stop if it runs more than 30 minutes (should be ~1-2 min for ~70 rows).
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::FromMinutes(30)) `
    -MultipleInstances IgnoreNew

# Principal: SYSTEM (or specified user). Service Account would be:
#   -UserId 'AraSVendorSync' -LogonType ServiceAccount
$principal = New-ScheduledTaskPrincipal -UserId $User -LogonType ServiceAccount -RunLevel Highest

# Register (overwrite if present)
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Removing existing task '$TaskName'..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Register-ScheduledTask -TaskName $TaskName `
    -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description 'Refresh per-vendor pricing on Aras Vendor Part rows from Digi-Key/Mouser APIs. See https://github.com/Tungsten-lawrenceb/aras-vendor-sync.' `
    | Out-Null

Write-Host "Registered scheduled task '$TaskName'"
Write-Host "  trigger: $($DayOfWeek -join ',') at $At"
Write-Host "  runs as: $User"
Write-Host "  script:  $ScriptPath"
Write-Host "  config:  $ConfigPath"
Write-Host ""
Write-Host "Tip: run it once manually first:"
Write-Host "  Start-ScheduledTask -TaskName $TaskName"
Write-Host "  Get-ScheduledTaskInfo -TaskName $TaskName"
