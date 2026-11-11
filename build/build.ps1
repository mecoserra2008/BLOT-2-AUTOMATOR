<#
.SYNOPSIS
  Builds dist\BLOT_BestEx.xlsm from src\, runs the VBA self-tests headless and cross-checks
  them with tests\reference_impl.py. Fails (exit 1) on any compile / test / cross-check error.

.PARAMETER SkipReference
  Do not run the Python reference cross-check (not recommended).

.PARAMETER KeepVbomSetting
  Leave "Trust access to the VBA project object model" enabled after the build
  (by default the previous registry value is restored).

.NOTES
  Needs Excel and "Trust access to the VBA project object model". The script enables it for the
  current user (HKCU ...\Excel\Security\AccessVBOM = 1) during the build and restores it afterwards.
  If a group policy blocks it, use the manual assembly steps printed at the end.
#>
[CmdletBinding()]
param([switch]$SkipReference, [switch]$KeepVbomSetting)

$ErrorActionPreference = 'Stop'
$root    = Split-Path $PSScriptRoot -Parent
$srcDir  = Join-Path $root 'src'
$distDir = Join-Path $root 'dist'
$out     = Join-Path $distDir 'BLOT_BestEx.xlsm'
$fixDir  = Join-Path $root 'tests\fixtures'
$vbaOut  = Join-Path $root 'tests\vba_out'
$modules = 'modConfig.bas', 'modBlotBestEx.bas', 'modSelfTest.bas'
$secKey  = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$polKey  = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security'

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

$manual = @"
MANUAL ASSEMBLY (if VBA project access is blocked):
  1. New workbook, Alt+F11.
  2. File > Import File: src\modConfig.bas, src\modBlotBestEx.bas, src\modSelfTest.bas.
  3. Paste src\ThisWorkbook.vba into the ThisWorkbook module (named 'EstaPasta_de_trabalho' in PT Excel).
  4. Run Setup, then EmbedFixtures with the path of tests\fixtures (Immediate window:
     EmbedFixtures "<repo>\tests\fixtures"), then RunSelfTests.
  5. Save as dist\BLOT_BestEx.xlsm (Excel Macro-Enabled Workbook).
"@

# ---------------------------------------------------------------- 1. sources: ASCII + CRLF
Write-Step 'Checking sources (ASCII only, CRLF line endings)'
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('blot_build_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
foreach ($f in $modules + 'ThisWorkbook.vba') {
    $p = Join-Path $srcDir $f
    if (-not (Test-Path $p)) { throw "Missing source file: $p" }
    $bytes = [IO.File]::ReadAllBytes($p)
    $bad = 0; foreach ($b in $bytes) { if ($b -gt 127) { $bad++ } }
    if ($bad -gt 0) { throw "$f contains $bad non-ASCII byte(s): the VBA import would corrupt them." }
    $text = [Text.Encoding]::ASCII.GetString($bytes) -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    [IO.File]::WriteAllText((Join-Path $tmp $f), $text, [Text.Encoding]::ASCII)
}

# ---------------------------------------------------------------- 2. VBA project access
Write-Step 'Enabling programmatic access to the VBA project (current user)'
$pol = (Get-ItemProperty -Path $polKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
if ($null -ne $pol -and $pol -eq 0) { Write-Host $manual; throw 'AccessVBOM is disabled by group policy.' }
if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
$prevVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -Type DWord

$xl = $null; $wb = $null; $exit = 0; $buildPid = @()
try {
    # ------------------------------------------------------------ 3. assemble the workbook
    Write-Step 'Starting Excel'
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $xl = New-Object -ComObject Excel.Application
    $buildPid = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1          # msoAutomationSecurityLow: macros enabled for this session
    $wb  = $xl.Workbooks.Add()
    try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count }
    catch { Write-Host $manual; throw "Cannot access the VBA project ($($_.Exception.Message))." }

    Write-Step 'Importing modules'
    foreach ($m in $modules) { $null = $vbp.VBComponents.Import((Join-Path $tmp $m)) }

    $cn = $wb.CodeName
    $tw = $null
    if ($cn) { try { $tw = $vbp.VBComponents.Item($cn) } catch { $tw = $null } }
    if (-not $tw) {                                       # localized-safe fallback
        $sheetCodeNames = @(); foreach ($s in $wb.Worksheets) { $sheetCodeNames += $s.CodeName }
        foreach ($c in $vbp.VBComponents) { if ($c.Type -eq 100 -and $sheetCodeNames -notcontains $c.Name) { $tw = $c; break } }
    }
    if (-not $tw) { throw 'ThisWorkbook component not found.' }
    $twCode = [IO.File]::ReadAllText((Join-Path $tmp 'ThisWorkbook.vba'))
    $twCode = ($twCode -split "`r`n" | Where-Object { $_ -notmatch '^\s*Option Explicit\s*$' }) -join "`r`n"
    if ($tw.CodeModule.CountOfLines -gt 0) { $tw.CodeModule.DeleteLines(1, $tw.CodeModule.CountOfLines) }
    $tw.CodeModule.AddFromString("Option Explicit`r`n" + $twCode)
    Write-Host "    ThisWorkbook component: $($tw.Name)"

    $q = "'" + $wb.Name + "'!"
    Write-Step 'Running Setup'
    $null = $xl.Run($q + 'Setup')
    Write-Step 'Embedding fixtures'
    $null = $xl.Run($q + 'EmbedFixtures', $fixDir)

    # ------------------------------------------------------------ 4. self-tests
    Write-Step 'Running self-tests (headless)'
    $ok = $xl.Run($q + 'RunSelfTestsHeadless')
    $ts = $wb.Worksheets.Item('TESTS')
    Write-Host ('    ' + $ts.Range('B2').Text + ' | fixtures: ' + $ts.Range('D2').Text)
    $nCsv = @(Get-ChildItem $fixDir -Filter *.csv).Count
    if ([int]$ts.Range('D2').Value2 -ne $nCsv) { throw "Self-tests ran $($ts.Range('D2').Text) fixture(s), expected $nCsv." }
    if (-not $ok) {
        $last = $ts.Cells($ts.Rows.Count, 1).End(-4162).Row
        for ($r = 5; $r -le $last; $r++) {
            if ($ts.Cells($r, 6).Text -ne 'PASS') {
                Write-Host ('    FAIL {0} / {1} / {2}: expected [{3}] got [{4}]' -f $ts.Cells($r,1).Text, $ts.Cells($r,2).Text, $ts.Cells($r,3).Text, $ts.Cells($r,4).Text, $ts.Cells($r,5).Text) -ForegroundColor Red
            }
        }
        throw 'Self-tests failed - see TESTS sheet.'
    }
    if (Test-Path $vbaOut) { Remove-Item $vbaOut -Recurse -Force }
    $null = $xl.Run($q + 'DumpSelfTestResults', $vbaOut)

    # ------------------------------------------------------------ 5. reference cross-check
    if (-not $SkipReference) {
        Write-Step 'Cross-checking with tests\reference_impl.py'
        $py = Get-Command python -ErrorAction SilentlyContinue
        if (-not $py) { $py = Get-Command py -ErrorAction SilentlyContinue }
        if (-not $py) { throw 'Python not found (needed for the reference cross-check; use -SkipReference to bypass).' }
        & $py.Source (Join-Path $root 'tests\reference_impl.py') --vba-dir $vbaOut
        if ($LASTEXITCODE -ne 0) { throw 'Reference implementation does not match the VBA results.' }
    }

    # ------------------------------------------------------------ 6. save
    Write-Step "Saving $out"
    if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.Worksheets.Item('DASHBOARD').Activate()
    $wb.SaveAs($out, 52)                                  # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false); $wb = $null

    # ------------------------------------------------------------ 7. verify the saved file
    Write-Step 'Verifying the saved workbook'
    $xl.EnableEvents = $false                             # do not arm the scheduler here
    $wb = $xl.Workbooks.Open($out, 0, $true)
    $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }
    foreach ($need in 'README','DASHBOARD','CONFIG','DEALER_MAP','DEALS','HIST','HIST_DAILY','RAW','LOG','TESTS','FIXTURES') {
        if ($names -notcontains $need) { throw "Sheet $need missing in the saved workbook." }
    }
    $btn = @(); foreach ($sh in $wb.Worksheets.Item('DASHBOARD').Shapes) { if ($sh.Name -like 'btn_*') { $btn += $sh.Name } }
    if ($btn.Count -ne 8) { throw "DASHBOARD has $($btn.Count) buttons, expected 8." }
    $comps = @(); foreach ($c in $wb.VBProject.VBComponents) { $comps += $c.Name }
    foreach ($m in 'modConfig','modBlotBestEx','modSelfTest') { if ($comps -notcontains $m) { throw "Module $m missing." } }
    $ok2 = $xl.Run("'" + $wb.Name + "'!RunSelfTestsHeadless")
    if (-not $ok2) { throw 'Self-tests fail when re-run from the saved workbook.' }
    Write-Host "    Sheets: $($names -join ', ')"
    Write-Host "    Buttons: $($btn -join ', ')"
    $wb.Close($false); $wb = $null
    Write-Host ''
    Write-Host "BUILD OK: $out" -ForegroundColor Green
}
catch {
    Write-Host "BUILD FAILED: $($_.Exception.Message)" -ForegroundColor Red
    $exit = 1
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) {
        try { $xl.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Excel writes its trust settings back on exit: restore only after the build instance is gone
    foreach ($id in $buildPid) {
        $p = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(10000)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue; $null = $p.WaitForExit(5000) }
    }
    if (-not $KeepVbomSetting) {
        if ($null -eq $prevVbom) { Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $prevVbom -Type DWord }
        $now = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
        Write-Host "    AccessVBOM restored to: $(if ($null -eq $now) { '(not set)' } else { $now })"
    }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
exit $exit
