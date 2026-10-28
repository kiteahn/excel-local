$ErrorActionPreference = 'Stop'
$runtimeRoot = Join-Path $PSScriptRoot 'runtime'
$archivePath = Join-Path $runtimeRoot 'ollama-windows-amd64-v0.35.1.zip'
$binaryRoot = Join-Path $runtimeRoot 'ollama'
$expected = 'dc50b9ca7f9023c86525012632cd1615b093d0407987444a7f62ecab617e8e93'
New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
if (-not (Test-Path -LiteralPath (Join-Path $binaryRoot 'ollama.exe'))) {
    if (-not (Test-Path -LiteralPath $archivePath)) {
        curl.exe --fail --location --retry 2 --output $archivePath 'https://github.com/ollama/ollama/releases/download/v0.35.1/ollama-windows-amd64.zip'
        if ($LASTEXITCODE -ne 0) { throw 'Runtime download failed' }
    }
    if ((Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expected) { throw 'Runtime SHA-256 mismatch' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($archivePath)
    $resolvedRoot = [IO.Path]::GetFullPath($binaryRoot)
    foreach ($entry in $archive.Entries) {
        $destination = [IO.Path]::GetFullPath([IO.Path]::Combine($resolvedRoot,$entry.FullName))
        if (-not $destination.StartsWith($resolvedRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { $archive.Dispose(); throw 'Unsafe archive path' }
    }
    $archive.Dispose()
    Expand-Archive -LiteralPath $archivePath -DestinationPath $binaryRoot -Force
}
$env:OLLAMA_NO_CLOUD = '1'
$env:OLLAMA_HOST = '127.0.0.1:11435'
$env:OLLAMA_MODELS = Join-Path $runtimeRoot 'models'
$env:OLLAMA_CONTEXT_LENGTH = '4096'
$env:OLLAMA_MAX_LOADED_MODELS = '1'
$env:OLLAMA_NUM_PARALLEL = '1'
try { Invoke-RestMethod 'http://127.0.0.1:11435/api/version' -TimeoutSec 2 | Out-Null } catch {
    $runtimeProcess = Start-Process -FilePath (Join-Path $binaryRoot 'ollama.exe') -ArgumentList 'serve' -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runtimeRoot 'server-out.log') -RedirectStandardError (Join-Path $runtimeRoot 'server.log')
    $runtimeProcess.Id | Set-Content -LiteralPath (Join-Path $runtimeRoot 'server.pid')
    $ready = $false
    for ($i=0;$i -lt 40;$i++) {
        try { Invoke-RestMethod 'http://127.0.0.1:11435/api/version' -TimeoutSec 1 | Out-Null; $ready=$true; break } catch { Start-Sleep -Milliseconds 250 }
    }
    if (-not $ready) { throw 'Ollama did not start' }
}
& (Join-Path $binaryRoot 'ollama.exe') pull 'qwen3:4b-q4_K_M'
if ($LASTEXITCODE -ne 0) { throw 'Model download failed; no cloud fallback' }
& (Join-Path $binaryRoot 'ollama.exe') list
$modelInfo = (Invoke-RestMethod 'http://127.0.0.1:11435/api/tags').models | Where-Object name -eq 'qwen3:4b-q4_K_M'
if ($modelInfo.digest -ne '2bfd38a7daaf4b1037efe517ccb73d1a3bbd4822cf89f1a82be1569050a114e0') { throw 'Unexpected model digest; review before accepting a changed model' }
@{version='0.35.1'; model=$modelInfo.name; model_digest=$modelInfo.digest; binary_sha256=(Get-FileHash -LiteralPath (Join-Path $binaryRoot 'ollama.exe') -Algorithm SHA256).Hash.ToLowerInvariant(); prepared=[DateTime]::Now.ToString('o')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runtimeRoot 'runtime-lock.json') -Encoding UTF8
