# Chạy Muse lặp lại cho đến khi có file STOP.txt hoặc hết số lượt tối đa.
# Cách dùng (PowerShell, trong thư mục repo):
#   .\run-muse-loop.ps1                 # mặc định tối đa 30 lượt
#   .\run-muse-loop.ps1 -MaxRuns 100
# Muốn dừng: tạo file STOP.txt trong thư mục repo (hoặc nhấn Ctrl+C).
param(
  [int]$MaxRuns = 30,
  [string]$PromptFile = "PROMPT_MUSE.md",
  [int]$PauseSeconds = 15,
  # Điền thêm cờ cho muse exec nếu cần (kiểm tra bằng: muse exec --help)
  [string[]]$MuseArgs = @()
)

Set-Location $PSScriptRoot
if (-not (Test-Path $PromptFile)) { Write-Error "Không thấy $PromptFile"; exit 1 }
if (-not (Get-Command muse -ErrorAction SilentlyContinue)) { Write-Error "Không tìm thấy lệnh 'muse'. Cài và chạy 'muse login' trước."; exit 1 }
if (-not (Test-Path ".git")) { Write-Error "Thư mục này không phải repo git."; exit 1 }

$logDir = ".work\muse-logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

# Luôn làm trên nhánh riêng
$branch = (git branch --show-current).Trim()
if ($branch -ne "muse-upgrade") {
  git checkout -B muse-upgrade | Out-Null
}

for ($i = 1; $i -le $MaxRuns; $i++) {
  if (Test-Path "STOP.txt") { Write-Host "Thấy STOP.txt - dừng."; break }
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  Write-Host "=== Lượt $i/$MaxRuns  $stamp ==="
  muse exec --prompt-file $PromptFile @MuseArgs 2>&1 | Tee-Object -FilePath "$logDir\run-$i-$stamp.log"

  # Lưu điểm an toàn nếu Muse để lại thay đổi chưa commit
  git add -A | Out-Null
  git diff --cached --quiet
  if ($LASTEXITCODE -ne 0) { git commit -m "muse run $i checkpoint" | Out-Null }

  if (Test-Path "STOP.txt") { Write-Host "Thấy STOP.txt - dừng."; break }
  Start-Sleep -Seconds $PauseSeconds
}
Write-Host "Kết thúc vòng lặp. Xem UPGRADE_LOG.md và NEXT.md."
