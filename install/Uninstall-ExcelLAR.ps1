<#
.SYNOPSIS
    Removes the ExcelLAR add-in for the current Windows user. OpenSolver is left in place.
.PARAMETER Target
    Folder where the add-in was installed. Default: the Excel add-in folder of the user.
.PARAMETER Yes
    Answers yes to every question.
#>
[CmdletBinding()]
param(
    [string]$Target = (Join-Path $env:APPDATA 'Microsoft\AddIns'),
    [switch]$Yes
)
$ErrorActionPreference = 'Stop'
$dst = Join-Path $Target 'ExcelLAR.xlam'
Write-Host 'ExcelLAR uninstaller' -ForegroundColor White
while (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    if ($Yes) { Write-Host 'Excel must be closed first.' -ForegroundColor Red; exit 1 }
    $a = Read-Host 'Excel is open. Close every Excel window, then press Enter (Q to quit)'
    if ($a -match '^[Qq]') { exit 1 }
}
$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    foreach ($ai in @($xl.AddIns)) {
        if ($ai.FullName -eq $dst -and $ai.Installed) { $ai.Installed = $false; Write-Host "Deactivated $dst" -ForegroundColor Green }
    }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($xl) { try { $xl.Quit() } catch { } }
    foreach ($o in @($wb, $xl)) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if (Test-Path $dst) {
    $ok = $Yes
    if (-not $ok) { $a = Read-Host "Delete $dst ? [Y/n]"; $ok = ($a -eq '' -or $a -match '^[Yy]') }
    if ($ok) { Remove-Item $dst -Force; Write-Host "Deleted $dst" -ForegroundColor Green }
}
Write-Host 'ExcelLAR has been removed. OpenSolver was not changed.'
exit 0
