<#
.SYNOPSIS
    Register the Windows Scheduled Task that runs the datasheet-queue
    drainer every 15 minutes.

.PARAMETER ScriptPath
    Path to Drain-DatasheetQueue.ps1.

.PARAMETER ConfigPath
    Path to the JSON config. Defaults to
    C:\ProgramData\AarasVendorSync\config.json.

.PARAMETER TaskName
    Scheduled Task name. Defaults to 'aras-datasheet-drain'.

.PARAMETER IntervalMinutes
    Polling interval. Defaults to 15.

.PARAMETER User
    Identity the task runs as. Defaults to SYSTEM.
#>

[CmdletBinding()]
param(
    [string]$ScriptPath,
    [string]$ConfigPath = "$env:ProgramData\AarasVendorSync\config.json",
    [string]$TaskName   = 'aras-datasheet-drain',
    [int]$IntervalMinutes = 15,
    [string]$User       = 'SYSTEM'
)

$ErrorActionPreference = 'Stop'

if (-not $ScriptPath) {
    $here = $PSScriptRoot
    if (-not $here) { $here = Split-Path $MyInvocation.MyCommand.Path -Parent }
    $ScriptPath = Join-Path (Split-Path $here -Parent) 'src\Drain-DatasheetQueue.ps1'
}

if (-not (Test-Path $ScriptPath)) { throw "Script not found: $ScriptPath" }

$psExe = Get-Command pwsh -ErrorAction SilentlyContinue
if (-not $psExe) { $psExe = Get-Command powershell }

$action = New-ScheduledTaskAction `
    -Execute $psExe.Source `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -ConfigPath `"$ConfigPath`""

# Trigger: every $IntervalMinutes minutes, starting now, repeating indefinitely.
$now = Get-Date
$trigger = New-ScheduledTaskTrigger -Once -At $now `
    -RepetitionInterval ([TimeSpan]::FromMinutes($IntervalMinutes))

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::FromMinutes(10)) `
    -MultipleInstances IgnoreNew

$principal = New-ScheduledTaskPrincipal -UserId $User -LogonType ServiceAccount -RunLevel Highest

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Removing existing task '$TaskName'..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}
Register-ScheduledTask -TaskName $TaskName `
    -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description 'Drain Aras DatasheetFetchRequest queue every 15 min. Downloads datasheets from Digi-Key and attaches them to Manufacturer Parts. See https://github.com/Tungsten-lawrenceb/aras-vendor-sync.' `
    | Out-Null

Write-Host "Registered scheduled task '$TaskName'"
Write-Host "  trigger: every $IntervalMinutes min starting $now"
Write-Host "  runs as: $User"
Write-Host "  script:  $ScriptPath"
