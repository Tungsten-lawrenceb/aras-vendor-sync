<#
.SYNOPSIS
    Remove the aras-vendor-sync Scheduled Task.

.PARAMETER TaskName
    Defaults to 'aras-vendor-sync'.
#>
[CmdletBinding()]
param([string]$TaskName = 'aras-vendor-sync')

$ErrorActionPreference = 'Stop'
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Unregistered task '$TaskName'."
} else {
    Write-Host "Task '$TaskName' was not registered; nothing to do."
}
