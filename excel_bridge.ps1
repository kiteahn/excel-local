param([Parameter(Mandatory=$true)][string]$JobFile,[switch]$CleanupOnly)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if ($CleanupOnly) {
    $recordFile = Join-Path (Split-Path -Parent $JobFile) 'excel-process.json'
    if (Test-Path -LiteralPath $recordFile) {
        $record = Get-Content -LiteralPath $recordFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($record.owned -eq $true) {
            $process = Get-Process -Id ([int]$record.pid) -ErrorAction SilentlyContinue
            if ($null -ne $process -and $process.ProcessName -eq 'EXCEL' -and $process.Path -eq $record.path -and $process.StartTime.ToUniversalTime() -eq [DateTime]::Parse($record.started).ToUniversalTime()) {
                Stop-Process -Id $process.Id -Force
                if (-not $process.WaitForExit(10000)) { [Environment]::Exit(1) }
            }
        }
    }
    [Environment]::Exit(0)
}
$job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = $null
$book = $null
$result = $null
$scriptStartUtc = [DateTime]::UtcNow
$ownedExcelStarted = $null
$ownsExcel = $false
$clock = [Diagnostics.Stopwatch]::StartNew()
$timings = @{}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelLocalProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@

function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function Get-PatchRuns($patches) {
    $runs = [Collections.Generic.List[object]]::new()
    $current = $null
    foreach ($patch in ($patches | Sort-Object col,row)) {
        if ($null -eq $current -or $patch.col -ne $current.col -or $patch.row -ne ($current.start + $current.patches.Count) -or $patch.format -cne $current.format) {
            $current = @{col=[int]$patch.col; start=[int]$patch.row; format=$patch.format; patches=[Collections.Generic.List[object]]::new()}
            $runs.Add($current)
        }
        $current.patches.Add($patch)
    }
    return $runs.ToArray()
}

function Assert-NoErrors($sheet) {
    foreach ($cellKind in @(2,-4123)) {
        $errors = $null
        try { $errors = $sheet.UsedRange.SpecialCells($cellKind,16) }
        catch [Runtime.InteropServices.COMException] {
            # Excel raises 1004 when SpecialCells finds no matching cells.
            if ($_.Exception.HResult -ne -2146827284) { throw }
        }
        if ($null -ne $errors) { Release-Com $errors; throw 'Worksheet contains Excel error values after recalculation' }
    }
}

function Write-Grid($sheet, $rows) {
    $height = $rows.Count
    $width = $rows[0].Count
    $matrix = [object[,]]::new($height, $width)
    for ($r = 0; $r -lt $height; $r++) {
        for ($c = 0; $c -lt $width; $c++) {
            $value = $rows[$r][$c]
            if ($value -is [long] -or $value -is [decimal] -or $value -is [int]) {
                $value = [double]$value
            }
            $matrix[$r,$c] = $value
        }
    }
    $range = $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($height,$width))
    # Text format is set BEFORE assignment: imported strings never become formulas.
    $range.NumberFormat = '@'
    $range.Value2 = $matrix
    for ($c = 0; $c -lt $width; $c++) {
        $allNumeric = $true
        for ($r = 1; $r -lt $height; $r++) {
            if ($null -ne $rows[$r][$c] -and $rows[$r][$c] -isnot [ValueType]) {
                $allNumeric = $false
                break
            }
        }
        if ($allNumeric -and $height -gt 1) {
            $sheet.Range($sheet.Cells.Item(2,$c+1),$sheet.Cells.Item($height,$c+1)).NumberFormat = 'General'
        }
    }
    $header = $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item(1,$width))
    $header.Font.Bold = $true
    $header.Interior.Color = 15128749
    $range.Font.Name = 'Calibri'
    $range.Font.Size = 11
    [void]$range.Columns.AutoFit()
    for ($c = 1; $c -le $width; $c++) {
        if ($sheet.Columns.Item($c).ColumnWidth -gt 42) { $sheet.Columns.Item($c).ColumnWidth = 42 }
    }
    $range.WrapText = $true
    [void]$range.Rows.AutoFit()
    $sheet.PageSetup.Orientation = 2
    $sheet.PageSetup.PaperSize = 9
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = $false
    $sheet.PageSetup.PrintTitleRows = '$1:$1'
    $sheet.PageSetup.PrintArea = $range.Address()
    Release-Com $header
    Release-Com $range
}

try {
    if ($job.operation -notin @('inspect','write')) { throw 'Unsupported operation' }
    $preExistingExcelIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcessId = 0
    [void][ExcelLocalProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
    $excelProcess = Get-Process -Id $excelProcessId
    $ownedExcelStarted = $excelProcess.StartTime.ToUniversalTime()
    $ownsExcel = $excelProcessId -notin $preExistingExcelIds -and $ownedExcelStarted -ge $scriptStartUtc.AddSeconds(-3)
    @{pid=$excelProcessId; started=$ownedExcelStarted.ToString('o'); owned=$ownsExcel; path=$excelProcess.Path} | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $JobFile) 'excel-process.json') -Encoding UTF8
    if (-not $ownsExcel) { throw 'Excel did not create an isolated instance; existing Excel session was not changed' }
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    if ($job.input) {
        # Inputs have already passed OOXML preflight, including connections.
        $book = $excel.Workbooks.Open([string]$job.input, 0, $true)
        if ($book.Connections.Count -gt 0) { throw 'Workbook has data connections' }
    } else {
        $book = $excel.Workbooks.Add()
    }
    $originalCalculation = $excel.Calculation
    $excel.Calculation = -4135
    $timings.open_seconds = $clock.Elapsed.TotalSeconds

    if ($job.operation -eq 'inspect') {
        $excel.CalculateFullRebuild()
        if ($excel.CalculationState -ne 0) { throw 'Excel calculation did not finish' }
        $names = @()
        foreach ($tab in $book.Worksheets) { $names += [string]$tab.Name; Release-Com $tab }
        $name = if ($job.sheet) { [string]$job.sheet } else { $names[0] }
        $sheet = $book.Worksheets.Item($name)
        if (-not $job.lenient) { Assert-NoErrors $sheet }
        $lastRowCell = $sheet.Cells.Find('*', [Type]::Missing, -4123, 2, 1, 2, $false, $false, $false)
        $lastColCell = $sheet.Cells.Find('*', [Type]::Missing, -4123, 2, 2, 2, $false, $false, $false)
        if ($null -eq $lastRowCell) { throw 'Worksheet is empty' }
        $height = [int]$lastRowCell.Row
        $width = [int]$lastColCell.Column
        if ($height -gt 20001 -or $width -gt 60 -or ($height * $width) -gt 400000) {
            throw 'V1 limit: 20000 data rows, 60 columns, 400000 cells'
        }
        $range = $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item($height,$width))
        $matrix = $range.Value2
        $rows = [System.Collections.Generic.List[object]]::new()
        for ($r = 1; $r -le $height; $r++) {
            $row = [object[]]::new($width)
            for ($c = 1; $c -le $width; $c++) {
                $row[$c-1] = if ($height -eq 1 -and $width -eq 1) { $matrix } else { $matrix[$r,$c] }
            }
            $rows.Add($row)
        }
        $result = @{ ok=$true; sheets=$names; sheet=$name; rows=$rows.ToArray(); excel_version=[string]$excel.Version }
        Release-Com $lastRowCell
        Release-Com $lastColCell
        Release-Com $range
        Release-Com $sheet
    } else {
        $phase = $clock.Elapsed.TotalSeconds
        foreach ($spec in $job.sheets) {
            $sheet = $book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))
            $sheet.Name = [string]$spec.name
            Write-Grid $sheet $spec.rows
            foreach ($run in (Get-PatchRuns $spec.formulas)) {
                $range = $sheet.Range($sheet.Cells.Item($run.start,$run.col),$sheet.Cells.Item(($run.start+$run.patches.Count-1),$run.col))
                $matrix = [object[,]]::new($run.patches.Count,1)
                for ($r=0; $r -lt $run.patches.Count; $r++) { $matrix[$r,0] = [string]$run.patches[$r].formula }
                $range.NumberFormat = 'General'
                $range.Formula = $matrix
                Release-Com $range
            }
            foreach ($run in (Get-PatchRuns $spec.formats)) {
                $range = $sheet.Range($sheet.Cells.Item($run.start,$run.col),$sheet.Cells.Item(($run.start+$run.patches.Count-1),$run.col))
                $range.NumberFormat = [string]$run.format
                Release-Com $range
            }
            Release-Com $sheet
        }
        $timings.write_seconds = $clock.Elapsed.TotalSeconds - $phase
        $phase = $clock.Elapsed.TotalSeconds
        $excel.CalculateFullRebuild()
        if ($excel.CalculationState -ne 0) { throw 'Excel calculation did not finish' }
        $timings.calculate_seconds = $clock.Elapsed.TotalSeconds - $phase
        $phase = $clock.Elapsed.TotalSeconds
        $verified = 0
        foreach ($spec in $job.sheets) {
            $checkSheet = $book.Worksheets.Item([string]$spec.name)
            Assert-NoErrors $checkSheet
            foreach ($run in (Get-PatchRuns $spec.formulas)) {
                $range = $checkSheet.Range($checkSheet.Cells.Item($run.start,$run.col),$checkSheet.Cells.Item(($run.start+$run.patches.Count-1),$run.col))
                if ($range.HasFormula -ne $true) { throw 'Missing formula in result' }
                $values = $range.Value2
                for ($r=0; $r -lt $run.patches.Count; $r++) {
                $patch = $run.patches[$r]
                $actual = if ($run.patches.Count -eq 1) { $values } else { $values[($r+1),1] }
                if ($patch.expected -is [ValueType] -and $patch.expected -isnot [bool]) {
                    $tolerance = [double]$job.tolerance
                    if ($actual -isnot [ValueType] -or $actual -is [bool] -or [double]::IsNaN([double]$actual) -or [double]::IsInfinity([double]$actual) -or [Math]::Abs([double]$actual - [double]$patch.expected) -gt $tolerance) {
                        throw 'Independent numeric reconciliation failed'
                    }
                } elseif ($patch.expected -is [bool]) {
                    if ($actual -isnot [bool] -or $actual -ne $patch.expected) { throw 'Independent boolean reconciliation failed' }
                } elseif ([string]$actual -cne [string]$patch.expected) {
                    throw 'Independent value reconciliation failed'
                }
                $verified++
                }
                Release-Com $range
            }
            Release-Com $checkSheet
        }
        $timings.verify_seconds = $clock.Elapsed.TotalSeconds - $phase
        $phase = $clock.Elapsed.TotalSeconds
        $resultSheet = $book.Worksheets.Item([string]$job.result_sheet)
        [void]$resultSheet.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
        $book.CheckCompatibility = $false
        $excel.Calculation = $originalCalculation
        $book.SaveAs([string]$job.output,51)
        if ($job.pdf) { $resultSheet.ExportAsFixedFormat(0,[string]$job.pdf) }
        $timings.save_seconds = $clock.Elapsed.TotalSeconds - $phase
        $result = @{ok=$true; formula_checks=$verified; excel_version=[string]$excel.Version; result_sheet=$job.result_sheet; timings=$timings}
        Release-Com $resultSheet
    }
} catch {
    $result = @{ok=$false; error=$_.Exception.Message}
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch { }; Release-Com $book }
    if ($null -ne $excel) { if ($ownsExcel) { try { $excel.Quit() } catch { } }; Release-Com $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # COM reference wrappers can keep an invisible Excel instance alive after Quit.
    # Only clean up the PID of a NEW instance created by this invocation.
    if ($ownsExcel) {
        $ownedProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        if ($null -ne $ownedProcess -and $ownedProcess.ProcessName -eq 'EXCEL' -and $ownedProcess.StartTime.ToUniversalTime() -eq $ownedExcelStarted) {
            Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
        }
    }
}
$result | ConvertTo-Json -Depth 20 -Compress | Set-Content -LiteralPath $job.result_file -Encoding UTF8
if (-not $result.ok) { [Environment]::Exit(1) }
[Environment]::Exit(0)
