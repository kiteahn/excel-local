#Requires -Version 5.1
<#
    run_all_tests.ps1 — chạy toàn bộ lưới kiểm thử của Excel Local Expert trên Windows.

    Nhóm 1: không cần Excel, không cần màn hình.
    Nhóm 2: cần Windows (DPAPI thật), không cần Excel.
    Nhóm 3: cần Excel desktop thật (và màn hình cho GUI).

    Cuối lượt chạy, script kiểm tra:
      - thư mục examples/ không bị test ghi đè;
      - không còn tiến trình Excel do bridge tạo (đọc .work/*/excel-process.json);
      - không còn thư mục tạm mới do lượt chạy này để lại.

    Exit code 0 chỉ khi mọi lệnh exit 0, examples/ sạch và không có Excel mồ côi.
    Không push, không gọi cloud: đây là script kiểm thử cục bộ.
#>
[CmdletBinding()]
param(
    [switch]$SkipExcel,
    [switch]$SkipGui,
    [switch]$SkipBenchmark,
    [string[]]$Only
)

$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
if (-not $root) { $root = Split-Path -Parent $MyInvocation.MyCommand.Path }
Push-Location $root
$startedAt = Get-Date
$runId = $startedAt.ToString('yyyyMMdd-HHmmss')
$work = Join-Path $root '.work'
$logDir = Join-Path $work 'qa-logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir "run-$runId.log"
$transcriptRows = New-Object System.Collections.Generic.List[object]

function Write-Log([string]$text) {
    $line = "[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss'), $text
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Get-PythonCommand {
    $candidates = @('python', 'py')
    foreach ($candidate in $candidates) {
        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($command) { return $command.Source }
    }
    throw 'Không tìm thấy python trên PATH.'
}

function Get-ExamplesState {
    if (Get-Command git -ErrorAction SilentlyContinue) {
        return (git status --porcelain -- examples | Out-String).Trim()
    }
    $folder = Join-Path $root 'examples'
    if (-not (Test-Path $folder)) { return '' }
    return (Get-ChildItem -LiteralPath $folder -File | Sort-Object Name |
        ForEach-Object { "{0} {1}" -f $_.Name, (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join "`n"
}

function Get-OwnedExcelFingerprint {
    $folder = Join-Path $root '.work'
    if (-not (Test-Path $folder)) { return @{} }
    $map = @{}
    Get-ChildItem -LiteralPath $folder -Recurse -Filter 'excel-process.json' -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            $map[$_.FullName] = '{0}|{1}' -f $_.LastWriteTimeUtc.Ticks, (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        }
    return $map
}

function Get-LeftoverOwnedExcel([hashtable]$before) {
    $orphans = New-Object System.Collections.Generic.List[object]
    $after = Get-OwnedExcelFingerprint
    foreach ($path in $after.Keys) {
        if ($before.ContainsKey($path) -and $before[$path] -eq $after[$path]) { continue }
        try { $record = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
        if ($record.owned -ne $true -or -not $record.pid) { continue }
        $process = Get-Process -Id ([int]$record.pid) -ErrorAction SilentlyContinue
        if ($null -eq $process -or $process.ProcessName -ne 'EXCEL') { continue }
        $samePath = $true
        try { $samePath = ($process.Path -eq $record.path) } catch { $samePath = $false }
        $sameStart = $true
        try { $sameStart = ($process.StartTime.ToUniversalTime() -eq [DateTime]::Parse($record.started).ToUniversalTime()) } catch { $sameStart = $false }
        if ($samePath -and $sameStart) {
            $orphans.Add([pscustomobject]@{ Pid = $process.Id; Record = $path })
        }
    }
    return $orphans
}

function Get-WorkEntries {
    if (-not (Test-Path $work)) { return @() }
    return @(Get-ChildItem -LiteralPath $work -Force | Select-Object -ExpandProperty Name)
}

function Invoke-Suite([string]$group, [string]$name, [string[]]$arguments) {
    if ($Only -and $Only.Count -gt 0 -and -not ($Only | Where-Object { $name -like "*$_*" -or $_ -eq $group })) {
        return
    }
    $suiteLog = Join-Path $logDir ("{0}-{1}.log" -f $runId, ($name -replace '[^A-Za-z0-9_.-]', '_'))
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $output = & $script:python -X utf8 $name @arguments 2>&1
    $exit = $LASTEXITCODE
    $stopwatch.Stop()
    $output | Out-File -LiteralPath $suiteLog -Encoding UTF8
    $text = ($output | Out-String).Trim()
    $tail = ($text -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 1)
    $passed = $null
    $match = [regex]::Match($text, '"passed"\s*:\s*(\d+)')
    if ($match.Success) { $passed = [int]$match.Groups[1].Value }
    $row = [pscustomobject]@{
        Group   = $group
        Suite   = $name + $(if ($arguments.Count) { ' ' + ($arguments -join ' ') } else { '' })
        Exit    = $exit
        Seconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
        Checks  = $passed
        LastLine = $tail
    }
    $transcriptRows.Add($row)
    $state = if ($exit -eq 0) { 'PASS' } else { 'FAIL' }
    Write-Log ("{0,-6} {1} ({2}s){3} :: {4}" -f $state, $row.Suite, $row.Seconds, $(if ($null -ne $passed) { ", $passed check" } else { '' }), $tail)
}

$python = Get-PythonCommand
Write-Log "Bắt đầu run_all_tests.ps1 (log: $logPath)"
Write-Log "Python: $python"

$examplesBefore = Get-ExamplesState
$ownedBefore = Get-OwnedExcelFingerprint
$workBefore = Get-WorkEntries

$offline = @(
    @{ Name = 'test_app.py'; Args = @() },
    @{ Name = 'test_all_edge_cases.py'; Args = @() },
    @{ Name = 'test_upgrades.py'; Args = @() },
    @{ Name = 'test_cli.py'; Args = @() },
    @{ Name = 'test_headers.py'; Args = @() },
    @{ Name = 'test_usecases.py'; Args = @() },
    @{ Name = 'test_usecases2.py'; Args = @() },
    @{ Name = 'test_fuzz_light.py'; Args = @() },
    @{ Name = 'test_fuzz_headers.py'; Args = @() },
    @{ Name = 'test_fuzz_xlsx.py'; Args = @() },
    @{ Name = 'test_infer_golden.py'; Args = @() },
    @{ Name = 'test_sanitize_plan.py'; Args = @() },
    @{ Name = 'test_policy_defaults.py'; Args = @() },
    @{ Name = 'test_safety_limits.py'; Args = @() },
    @{ Name = 'test_prepare_sheets_golden.py'; Args = @() },
    @{ Name = 'test_batch_summary.py'; Args = @() },
    @{ Name = 'test_privacy_guard_cases.py'; Args = @() },
    @{ Name = 'test_single_file_filter.py'; Args = @() },
    @{ Name = 'test_diverse_scenarios_e2e.py'; Args = @() },
    @{ Name = 'test_batch_processing.py'; Args = @() },
    @{ Name = 'benchmark_python_stages.py'; Args = @() },
    @{ Name = 'secret_scan.py'; Args = @() }
)
$windows = @(
    @{ Name = 'test_router.py'; Args = @() },
    @{ Name = 'test_security_privacy.py'; Args = @() }
)
$excel = @(
    @{ Name = 'test_app.py'; Args = @('--integration') },
    @{ Name = 'test_read_boundary.py'; Args = @() }
)
if (-not $SkipGui) { $excel += @{ Name = 'test_gui.py'; Args = @() } }
if (-not $SkipBenchmark) { $excel += @{ Name = 'benchmark.py'; Args = @() } }

foreach ($suite in $offline) { Invoke-Suite '1-offline' $suite.Name $suite.Args }
foreach ($suite in $windows) { Invoke-Suite '2-windows' $suite.Name $suite.Args }
if ($SkipExcel) {
    Write-Log 'Bỏ qua nhóm 3 theo -SkipExcel.'
} else {
    foreach ($suite in $excel) { Invoke-Suite '3-excel' $suite.Name $suite.Args }
}

$examplesAfter = Get-ExamplesState
$orphans = @(Get-LeftoverOwnedExcel $ownedBefore)
$workAfter = Get-WorkEntries
$newWorkEntries = @($workAfter | Where-Object { $workBefore -notcontains $_ })

$failed = @($transcriptRows | Where-Object { $_.Exit -ne 0 })
$examplesDirty = ($examplesBefore -ne $examplesAfter)

Write-Log ''
Write-Log '===== BẢNG TỔNG HỢP ====='
$table = $transcriptRows | Format-Table -AutoSize Group, Suite, Exit, Seconds, Checks | Out-String
Write-Host $table
Add-Content -LiteralPath $logPath -Value $table -Encoding UTF8

if ($examplesDirty) {
    Write-Log 'FAIL: examples/ đã thay đổi trong lượt chạy này. git status --porcelain -- examples:'
    Write-Log ($examplesAfter | Out-String)
} else {
    Write-Log 'OK: examples/ không thay đổi.'
}
if ($orphans.Count) {
    foreach ($orphan in $orphans) { Write-Log ("FAIL: Excel mồ côi PID {0} (record {1})." -f $orphan.Pid, $orphan.Record) }
} else {
    Write-Log 'OK: không có Excel do bridge tạo còn sống.'
}
if ($newWorkEntries.Count) {
    Write-Log ("Thư mục tạm còn lại trong .work: {0}" -f ($newWorkEntries -join ', '))
} else {
    Write-Log 'OK: không có thư mục tạm mới còn lại trong .work.'
}

$exitCode = 0
if ($failed.Count -or $examplesDirty -or $orphans.Count) { $exitCode = 1 }
$totalSeconds = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 1)
Write-Log ("Kết thúc: {0} lệnh, {1} lỗi, {2}s. Exit {3}. Log: {4}" -f $transcriptRows.Count, $failed.Count, $totalSeconds, $exitCode, $logPath)
Pop-Location
exit $exitCode
