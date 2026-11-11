<#
.SYNOPSIS
    Installs the ExcelLAR add-in for the current Windows user and checks that it runs.

.DESCRIPTION
    The script performs the manual installation steps of INSTALL.md:
      1. checks that Microsoft Excel for Windows is installed and closed;
      2. verifies the SHA-256 checksum of src\ExcelLAR.xlam against SHA256SUMS.txt;
      3. copies the add-in to the Excel add-in folder of the user and removes the
         download mark of that copy only (the same as ticking "Unblock" in its Properties);
      4. finds OpenSolver, or opens its official download page and asks where it was extracted;
      5. registers both add-ins in Excel (File > Options > Add-ins);
      6. runs examples\Minimal_5Blocks.csv and compares the multiplier with the expected value.
    No Trust Center, macro or system setting is changed, and no administrator right is needed.
    A log is written to %TEMP%\ExcelLAR_install.log.

.PARAMETER Target
    Folder for the add-in. Default: the Excel add-in folder of the user.
.PARAMETER OpenSolverPath
    Full path of OpenSolver.xlam, if it is not found automatically.
.PARAMETER KeepOld
    Leaves an older ExcelLa/ExcelLAR add-in registered (by default it is deactivated, not deleted).
.PARAMETER SkipTest
    Skips the test run of step 6.
.PARAMETER Yes
    Answers yes to every question (unattended installation).
#>
[CmdletBinding()]
param(
    [string]$Target = (Join-Path $env:APPDATA 'Microsoft\AddIns'),
    [string]$OpenSolverPath = '',
    [switch]$KeepOld,
    [switch]$SkipTest,
    [switch]$Yes
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Log = Join-Path $env:TEMP 'ExcelLAR_install.log'
$ExpectedLambda = 1226.1996   # Minimal_5Blocks.csv, see results\Example_Minimal_5Blocks.xlsx
$OpenSolverPage = 'https://opensolver.org/installing-opensolver/'
try { Start-Transcript -Path $Log -Force | Out-Null } catch { }

function Step([string]$text) { Write-Host ''; Write-Host "== $text" -ForegroundColor Cyan }
function Ok([string]$text)   { Write-Host "   OK   $text" -ForegroundColor Green }
function Warn([string]$text) { Write-Host "   NOTE $text" -ForegroundColor Yellow }
function Fail([string]$text) {
    Write-Host "   STOP $text" -ForegroundColor Red
    Write-Host "   See INSTALL.md, section Troubleshooting. Log: $Log"
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}
function Ask([string]$question) {
    if ($Yes) { Write-Host "   $question [Y/n] y"; return $true }
    $a = Read-Host "   $question [Y/n]"
    return ($a -eq '' -or $a -match '^[YySs]')
}
function Release($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

Write-Host 'ExcelLAR 1.1 installer' -ForegroundColor White

# ---------------------------------------------------------------- 1. Excel
Step 'Checking Microsoft Excel'
if ($env:OS -ne 'Windows_NT') { Fail 'ExcelLAR needs Microsoft Excel for Windows.' }
if (-not [Type]::GetTypeFromProgID('Excel.Application')) { Fail 'Microsoft Excel for Windows was not found on this computer.' }
$platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
if ($platform) { Ok "Excel found (Microsoft 365, $platform)" } else { Ok 'Excel found' }
if ($platform -eq 'x86') { Warn 'This is 32-bit Excel. The code supports it, but only 64-bit Excel has been tested.' }
while (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Warn 'Excel is open. Save your work and close every Excel window.'
    if ($Yes) { Fail 'Excel must be closed before the installation.' }
    $a = Read-Host '   Press Enter when Excel is closed, or type Q to quit'
    if ($a -match '^[Qq]') { Fail 'Installation cancelled.' }
}
Ok 'Excel is closed'

# ---------------------------------------------------------------- 2. checksum
Step 'Checking the add-in file'
$src = Join-Path $Root 'src\ExcelLAR.xlam'
if (-not (Test-Path $src)) { Fail "File not found: $src. Extract the whole repository ZIP before running the installer." }
$sumFile = Join-Path $Root 'SHA256SUMS.txt'
if (Test-Path $sumFile) {
    $line = Select-String -Path $sumFile -Pattern 'ExcelLAR\.xlam' | Select-Object -First 1
    $expected = ($line.Line -split '\s+')[0].ToLower()
    $actual = (Get-FileHash $src -Algorithm SHA256).Hash.ToLower()
    if ($expected -ne $actual) { Fail 'The checksum of ExcelLAR.xlam does not match SHA256SUMS.txt. Download the repository again.' }
    Ok "SHA-256 matches ($($actual.Substring(0,12))...)"
} else { Warn 'SHA256SUMS.txt not found, checksum not verified.' }

# ---------------------------------------------------------------- 3. copy
Step "Copying the add-in to $Target"
New-Item -ItemType Directory -Force -Path $Target | Out-Null
$dst = Join-Path $Target 'ExcelLAR.xlam'
if (Test-Path $dst) {
    Copy-Item $dst "$dst.bak" -Force
    Warn "An earlier ExcelLAR.xlam was kept as $dst.bak"
}
Copy-Item $src $dst -Force
Unblock-File -Path $dst
Ok 'Add-in copied, and the download mark of this copy removed'

# ---------------------------------------------------------------- 4. OpenSolver
Step 'Looking for OpenSolver'
$os = $null
if ($OpenSolverPath) { if (Test-Path $OpenSolverPath) { $os = (Resolve-Path $OpenSolverPath).Path } else { Warn "Not found: $OpenSolverPath" } }
if (-not $os) {
    $opt = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options' -ErrorAction SilentlyContinue
    if ($opt) {
        foreach ($p in $opt.PSObject.Properties) {
            if ($p.Name -like 'OPEN*' -and "$($p.Value)" -match 'OpenSolver\.xlam') {
                $c = ("$($p.Value)" -replace '^/\w+\s+', '').Trim('"', ' ')
                if (-not [IO.Path]::IsPathRooted($c)) { $c = Join-Path $Target $c }
                if (Test-Path $c) { $os = $c; break }
            }
        }
    }
}
if (-not $os) {
    $places = @('C:\OpenSolver', $Target, (Join-Path $env:USERPROFILE 'Downloads'), (Join-Path $env:USERPROFILE 'Desktop'), (Join-Path $env:USERPROFILE 'Documents'))
    foreach ($pl in $places) {
        if (-not (Test-Path $pl)) { continue }
        $hit = Get-ChildItem -Path $pl -Filter 'OpenSolver.xlam' -Recurse -Depth 3 -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $os = $hit.FullName; break }
    }
}
if (-not $os) {
    Warn 'OpenSolver was not found. It is free software (GPL) distributed by its authors.'
    Write-Host '   1. Download the OpenSolver ZIP file for Windows from the page that opens now.'
    Write-Host '   2. Right-click the ZIP file, Properties, tick Unblock, OK.'
    Write-Host '   3. Extract it to a simple folder such as C:\OpenSolver (no accents or non-English characters).'
    if (-not $Yes) {
        Start-Process $OpenSolverPage
        $p = Read-Host '   Then paste here the full path of OpenSolver.xlam (Enter to skip)'
        if ($p) { $p = $p.Trim('"', ' '); if (Test-Path $p) { $os = (Resolve-Path $p).Path } else { Warn "Not found: $p" } }
    }
}
if ($os) {
    Ok "OpenSolver: $os"
    $blocked = Get-ChildItem (Split-Path $os) -Recurse -File -ErrorAction SilentlyContinue | Where-Object { Get-Item $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue }
    if ($blocked) {
        if (Ask "The OpenSolver files still carry the download mark ($(@($blocked).Count) files). Unblock them?") {
            $blocked | Unblock-File; Ok 'OpenSolver files unblocked'
        } else { Warn 'Excel may block OpenSolver until its files are unblocked.' }
    }
} else {
    Warn 'OpenSolver is not registered. ExcelLAR will open, but an analysis needs OpenSolver. Run this installer again after installing it.'
}

# ---------------------------------------------------------------- 5. register and 6. test
Step 'Registering the add-ins in Excel'
$xl = $null; $wb = $null; $book = $null
$result = 'not run'
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $ver = "$($xl.Version) build $($xl.Build)"
    Ok "Excel $ver started for the installation (hidden)"

    foreach ($ai in @($xl.AddIns)) {
        $n = "$($ai.Name)"
        if (($n -match '^ExcelLa(R)?\.xlam$') -and ($ai.FullName -ne $dst) -and $ai.Installed) {
            if ($KeepOld) { Warn "An older add-in stays active: $($ai.FullName). Both use Ctrl+Shift+A." }
            elseif (Ask "Deactivate the older add-in $($ai.FullName)? (the file is not deleted)") { $ai.Installed = $false; Ok "Deactivated $($ai.FullName)" }
        }
    }
    $paths = @($dst); if ($os) { $paths = @($os) + $paths }
    foreach ($p in $paths) {
        $ai = $xl.AddIns.Add($p, $false)
        if (-not $ai.Installed) { $ai.Installed = $true }
        $loaded = $false
        foreach ($b in @($xl.Workbooks)) { if ($b.FullName -eq $p) { $loaded = $true } }
        if (-not $loaded) { [void]$xl.Workbooks.Open($p) }
        Ok "Registered and loaded: $p"
    }

    if (-not $SkipTest) {
        Step 'Test run with examples\Minimal_5Blocks.csv'
        if (-not $os) { Warn 'Skipped: OpenSolver is missing.' }
        else {
            $csv = Join-Path $Root 'examples\Minimal_5Blocks.csv'
            $book = $xl.Workbooks.Open($csv)
            $sheet = $book.Worksheets.Item(1).Name
            $err = $xl.Run('ExcelLAR.xlam!RunAnalysisSilent', $sheet, $false)
            if ($err) { $result = "failed: $err" }
            else {
                $info = $book.Worksheets.Item('RunInfo')
                $lam = $null
                for ($r = 1; $r -le 30; $r++) { if ("$($info.Cells.Item($r,1).Value2)" -like 'lambda_u*') { $lam = [double]$info.Cells.Item($r,2).Value2 } }
                if ($null -ne $lam -and [math]::Abs($lam - $ExpectedLambda) -lt 1e-3 * $ExpectedLambda) {
                    $result = ('passed: lambda_u = {0:N4}, expected {1:N4}' -f $lam, $ExpectedLambda)
                } else { $result = "failed: lambda_u = $lam, expected $ExpectedLambda" }
            }
            if ($result -like 'passed*') { Ok "Test $result" } else { Warn "Test $result" }
        }
    }
}
catch { Warn "Excel automation error: $($_.Exception.Message)"; $result = 'error' }
finally {
    if ($book) { try { $book.Close($false) } catch { } }
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($xl) { try { $xl.Quit() } catch { } }
    Release $book; Release $wb; Release $xl
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Step 'Summary'
Write-Host "   Add-in:     $dst"
Write-Host "   OpenSolver: $(if ($os) { $os } else { 'not installed' })"
Write-Host "   Test:       $result"
Write-Host '   Open Excel and press Ctrl+Shift+A in any workbook to start ExcelLAR.'
Write-Host "   Log file:   $Log"
try { Stop-Transcript | Out-Null } catch { }
if ($result -like 'failed*' -or $result -eq 'error') { exit 2 }
exit 0
