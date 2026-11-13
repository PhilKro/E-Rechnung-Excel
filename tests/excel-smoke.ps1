<#
.SYNOPSIS
  End-to-end smoke test: loads the add-in into a hidden Excel instance and runs all five stages
  (create -> validate -> export -> import -> inbox) through the ERechnung_* macros in test mode (no dialogs).
  Keep this file ASCII-only: Windows PowerShell 5.1 reads BOM-less UTF-8 as ANSI, so umlauts are matched
  with "." in regexes and built with [char]0xFC in strings.

.DESCRIPTION
  Needs a desktop Excel (64-bit by default) and a built add-in (./build.ps1).
  KoSIT checks run only if the validator is installed (tools/setup-kosit.ps1).
  Leaves the produced files and a workbook (smoke.xlsx) in -OutDir for manual inspection.

.EXAMPLE
  ./tests/excel-smoke.ps1
  ./tests/excel-smoke.ps1 -Xll .\src\ERechnung.ExcelAddIn\bin\Release\net8.0-windows\publish\ERechnung-AddIn64-packed.xll
#>
param(
    [string]$Xll = (Join-Path $PSScriptRoot '..\src\ERechnung.ExcelAddIn\bin\Debug\net8.0-windows\publish\ERechnung-AddIn64-packed.xll'),
    [string]$OutDir = (Join-Path $env:TEMP ('ERechnungExcel-smoke\' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)

$ErrorActionPreference = 'Stop'
$source = (Resolve-Path $Xll).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
# Load a copy, so a hanging Excel instance never locks the build output.
$Xll = Join-Path $OutDir (Split-Path $source -Leaf)
Copy-Item $source $Xll

$failures = New-Object System.Collections.Generic.List[string]
function Check([bool]$condition, [string]$what) {
    if ($condition) { Write-Host "  OK   $what" -ForegroundColor Green }
    else { Write-Host "  FAIL $what" -ForegroundColor Red; $failures.Add($what) }
}

Add-Type -Namespace Smoke -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
'@

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
# Remember the PID of *this* Excel instance, so cleanup never touches other Excel windows of the user.
$excelPid = [uint32]0
[void][Smoke.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
$wb = $null
try {
    Write-Host "Loading $Xll"
    Check ([bool]$xl.RegisterXLL($Xll)) 'add-in loads (RegisterXLL)'
    $wb = $xl.Workbooks.Add()

    function Log { [string]$xl.Run('ERechnung_Testprotokoll') }

    # --- Stage 1+5: create a sheet, fill demo data -----------------------------------------
    Write-Host 'Stage 1: create invoice sheet with demo data'
    [void]$xl.Run('ERechnung_Testmodus', (Join-Path $OutDir 'demo_XRechnung.xml'), '')
    [void]$xl.Run('ERechnung_Beispiel')
    $ws = $xl.ActiveSheet
    Check ($ws.Name -eq 'Beispielrechnung') "demo sheet created (active: $($ws.Name))"
    Check ($ws.Range('ER_Number').Value2 -eq 'RE-2027-0042') 'invoice number written to named cell'
    Check ($ws.ListObjects.Count -eq 1 -and $ws.ListObjects.Item(1).ListRows.Count -eq 3) 'line table has 3 rows'
    $net = $ws.ListObjects.Item(1).ListColumns.Item('Nettobetrag').DataBodyRange.Value2
    $firstNet = $net.GetValue(1, 1)
    Check ([math]::Abs($firstNet - 510.0) -lt 0.001) "line net formula = 7.5 h x 68.00 = 510.00 (got $firstNet)"

    # A new, empty sheet: the pre-filled default rows of the line table must not count as lines
    [void]$xl.Run('ERechnung_Neu')
    $blank = $xl.ActiveSheet
    $blank.ListObjects.Item(1).ListColumns.Item('Bezeichnung').DataBodyRange.Cells.Item(1, 1).Value2 = 'Beratung'
    $blank.ListObjects.Item(1).ListColumns.Item('Menge').DataBodyRange.Cells.Item(1, 1).Value2 = 2
    $blank.ListObjects.Item(1).ListColumns.Item('Einzelpreis netto').DataBodyRange.Cells.Item(1, 1).Value2 = 100
    [void]$xl.Run('ERechnung_Pruefen')
    $log = Log
    Check ($log -notmatch 'BR-25' -and $blank.ListObjects.Item(1).HeaderRowRange.Interior.Color -ne (255 + 199*256 + 206*65536)) 'empty default rows are ignored'
    Check ($blank.Range('ER_Number').Interior.Color -eq (255 + 199*256 + 206*65536)) 'empty new sheet: missing invoice number highlighted'
    Remove-Variable blank
    $ws.Activate()

    # --- Stage 4: validate --------------------------------------------------------------------
    Write-Host 'Stage 4: validate'
    [void]$xl.Run('ERechnung_Pruefen')
    $log = Log
    Check ($log -match 'G.ltig') "check reports valid: $($log -replace '\s+',' ')"
    $totalsLabels = $ws.Range('E4:E26').Value2
    $row = 0; for ($i = 1; $i -le 23; $i++) { if ($totalsLabels.GetValue($i, 1) -eq 'Zahlbetrag') { $row = $i + 3 } }
    $due = if ($row) { $ws.Cells.Item($row, 6).Value2 } else { $null }
    # 510.00 + 749.70 = 1259.70 @19% -> 239.34; 79.90 @7% -> 5.59; gross 1584.53
    Check ($row -gt 0 -and [math]::Abs($due - 1584.53) -lt 0.001) "totals: Zahlbetrag = 1584.53 (got $due)"
    $reportName = "Pr$([char]0xFC)fbericht"
    $sheetNames = @($xl.ActiveWorkbook.Worksheets | ForEach-Object { $_.Name })
    Check ($sheetNames -contains $reportName) 'report sheet exists'

    # invalid input is flagged and highlighted
    $ws.Range('ER_BuyerReference').Value2 = $null
    [void]$xl.Run('ERechnung_Pruefen')
    $log = Log
    Check ($log -match 'Ung.ltig') 'missing Leitweg-ID makes the invoice invalid'
    Check ($ws.Range('ER_BuyerReference').Interior.Color -eq (255 + 199*256 + 206*65536)) 'missing field highlighted red'
    $report = $xl.ActiveWorkbook.Worksheets.Item($reportName)
    Check ($null -ne $report.UsedRange.Find('BR-DE-15')) 'report lists BR-DE-15'
    $ws.Range('ER_BuyerReference').Value2 = '04011000-12345-34'

    # --- Stage 2: export ----------------------------------------------------------------------
    Write-Host 'Stage 2: export'
    $ws.Activate()
    [void]$xl.Run('ERechnung_ExportUBL')
    $ubl = Join-Path $OutDir 'demo_XRechnung.xml'
    Check (Test-Path $ubl) 'XRechnung UBL written'
    $log = Log
    Check ($log -match 'Gespeichert' -and $log -notmatch 'ERROR') "UBL export message: $($log -replace '\s+',' ')"
    if (Test-Path $ubl) { Check ((Get-Content $ubl -Raw) -match 'xrechnung_3\.0') 'UBL has XRechnung 3.0 CustomizationID' }

    [void]$xl.Run('ERechnung_Testmodus', (Join-Path $OutDir 'demo_XRechnung_CII.xml'), '')
    $ws.Activate()
    [void]$xl.Run('ERechnung_ExportCII')
    Check (Test-Path (Join-Path $OutDir 'demo_XRechnung_CII.xml')) 'XRechnung CII written'
    [void](Log)

    [void]$xl.Run('ERechnung_Testmodus', (Join-Path $OutDir 'demo_ZUGFeRD.pdf'), '')
    $ws.Activate()
    [void]$xl.Run('ERechnung_ExportPDF')
    $pdf = Join-Path $OutDir 'demo_ZUGFeRD.pdf'
    Check (Test-Path $pdf) 'ZUGFeRD PDF written'
    $log = Log
    Check ($log -match 'Gespeichert' -and $log -notmatch 'ERROR') "PDF export message: $($log -replace '\s+',' ')"
    if (Test-Path $pdf) {
        $head = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($pdf), 0, 5)
        Check ($head -eq '%PDF-') 'PDF header'
    }

    # --- Stage 3: import ----------------------------------------------------------------------
    Write-Host 'Stage 3: import'
    [void]$xl.Run('ERechnung_Testmodus', '', $pdf)
    [void]$xl.Run('ERechnung_Import')
    $log = Log
    $in = $xl.ActiveSheet
    Check ($in.Name -like 'Eingang RE-2027-0042*') "received invoice sheet created (active: $($in.Name))"
    Check ($in.Range('ER_Seller_Name').Value2 -eq 'Musterbau GmbH') 'seller read from embedded XML'
    Check ($in.ListObjects.Item(1).ListRows.Count -eq 3) 'received invoice has 3 lines'
    Check ($log -match 'G.ltig') "received ZUGFeRD PDF is valid: $($log -replace '\s+',' ')"

    [void]$xl.Run('ERechnung_Testmodus', '', "$ubl|$(Join-Path $OutDir 'demo_XRechnung_CII.xml')")
    [void]$xl.Run('ERechnung_Import')
    $log = Log
    $inbox = $xl.ActiveWorkbook.Worksheets.Item('Rechnungseingang')
    $rows = $inbox.ListObjects.Item('ER_Eingang').ListRows.Count
    Check ($rows -eq 3) "inbox has 3 entries (got $rows)"
    Check ($log -match '2 E-Rechnung\(en\) eingelesen, davon 0 ung.ltig') "batch import summary: $($log -replace '\s+',' ')"

    # Stage 5 (Excel UI) is covered throughout: sheets, names, tables, drop-downs
    $wb.SaveAs((Join-Path $OutDir 'smoke.xlsx'), 51)
}
catch {
    Check $false "unexpected error: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
}
finally {
    try { if ($wb) { $wb.Close($false) } } catch { }
    try { $xl.Quit() } catch { }
    # Excel only exits once every COM reference held by this script is released.
    foreach ($name in 'ws', 'in', 'inbox', 'report', 'wb') { Remove-Variable -Name $name -ErrorAction SilentlyContinue }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
    Remove-Variable xl
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    # A hidden automation instance can linger while the host process lives; end exactly the one we started.
    if ($excelPid -ne 0) {
        $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($proc -and -not $proc.WaitForExit(5000)) { Stop-Process -Id $excelPid -Force }
    }
}

Write-Host ''
Write-Host "Output: $OutDir"
if ($failures.Count -gt 0) { Write-Host "$($failures.Count) check(s) failed" -ForegroundColor Red; exit 1 }
Write-Host 'All smoke checks passed' -ForegroundColor Green
