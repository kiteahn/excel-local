$ErrorActionPreference = 'Stop'
$workspace = (Resolve-Path -LiteralPath $PSScriptRoot).Path
$private = Join-Path $workspace ('.work\cleanup-'+[Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $private)
$jobFile = Join-Path $private 'job.json'
'{}' | Set-Content -LiteralPath $jobFile -Encoding UTF8
$excel = $null
$owned = $false
$checks = [Collections.Generic.List[string]]::new()
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CleanupTestProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
try {
    $before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $started = [DateTime]::UtcNow
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcessId = 0
    [void][CleanupTestProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelProcessId)
    $process = Get-Process -Id $excelProcessId
    $owned = $excelProcessId -notin $before -and $process.StartTime.ToUniversalTime() -ge $started.AddSeconds(-3)
    if (-not $owned) { throw 'Test could not acquire a new isolated Excel; existing session untouched' }
    $excel.Visible = $false
    $record = @{pid=$process.Id; started=$process.StartTime.ToUniversalTime().ToString('o'); path=$process.Path; owned=$false}
    $originalStarted = $record.started
    $originalPath = $record.path
    function Invoke-Cleanup {
        $record | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $private 'excel-process.json') -Encoding UTF8
        & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $workspace 'excel_bridge.ps1') -JobFile $jobFile -CleanupOnly
        if ($LASTEXITCODE -ne 0) { throw 'Cleanup invocation failed' }
    }
    Invoke-Cleanup
    if (-not (Get-Process -Id $process.Id -ErrorAction SilentlyContinue)) { throw 'Unowned process was stopped' }
    $checks.Add('Unowned Excel is not stopped')
    $record.owned = $true
    $record.started = $process.StartTime.ToUniversalTime().AddSeconds(1).ToString('o')
    Invoke-Cleanup
    if (-not (Get-Process -Id $process.Id -ErrorAction SilentlyContinue)) { throw 'Mismatched start time was accepted' }
    $checks.Add('Mismatched start time is not stopped')
    $record.started = $originalStarted
    $record.path = 'C:\not-the-owned-excel.exe'
    Invoke-Cleanup
    if (-not (Get-Process -Id $process.Id -ErrorAction SilentlyContinue)) { throw 'Mismatched executable was accepted' }
    $checks.Add('Mismatched executable path is not stopped')
    $record.path = $originalPath
    Invoke-Cleanup
    if (Get-Process -Id $process.Id -ErrorAction SilentlyContinue) { throw 'Owned Excel survived cleanup' }
    $checks.Add('Exact owned Excel is stopped')
    $result = @{passed=$checks.Count; checks=$checks.ToArray(); scope='Native isolated Excel; no workbook or user data'}
    $result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $workspace 'cleanup-test-results.json') -Encoding UTF8
    $result | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $excel) {
        if ($owned) { try { $excel.Quit() } catch { } }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if ($owned) {
        $remaining = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
        if ($null -ne $remaining -and $remaining.ProcessName -eq 'EXCEL' -and $remaining.Path -eq $originalPath -and $remaining.StartTime.ToUniversalTime() -eq [DateTime]::Parse($originalStarted).ToUniversalTime()) {
            Stop-Process -Id $remaining.Id -Force
        }
    }
    $resolvedPrivate = (Resolve-Path -LiteralPath $private).Path
    if (-not $resolvedPrivate.StartsWith((Join-Path $workspace '.work\'),[StringComparison]::OrdinalIgnoreCase)) { throw 'Temporary cleanup target outside workspace' }
    Remove-Item -LiteralPath $resolvedPrivate -Recurse -Force
}
