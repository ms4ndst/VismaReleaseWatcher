param(
    [int]$RestartDelaySeconds = 10
)

$ErrorActionPreference = 'Stop'
$AppName = 'VismaReleaseWatcher'
$ScriptPath = "C:\ProgramData\VismaReleaseWatcher\VismaReleaseWatcher.ps1"
$LogPath = "C:\ProgramData\VismaReleaseWatcher\watchdog.log"
$WinPS = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Write-Log($msg){
  try { "$((Get-Date).ToString('s')) $msg" | Add-Content -Path $LogPath -Encoding UTF8 } catch {}
}

if (-not (Test-Path -LiteralPath $ScriptPath)) {
  throw "Main script not found: $ScriptPath"
}

Write-Log "Watchdog starting for $ScriptPath"

while ($true) {
  # Use Windows PowerShell with a hidden window to avoid Windows Terminal flashes
  $p = Start-Process -FilePath $WinPS -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$ScriptPath`"" -WindowStyle Hidden -PassThru
  Write-Log "Started child PID=$($p.Id)"
  $p.WaitForExit()
  $exit = $p.ExitCode
  Write-Log "Child exited with code $exit"
  if ($exit -eq 200) { Write-Log "Intentional exit signaled. Watchdog stopping."; break }
  if ($exit -eq 1) {
    Write-Log "Crash exit detected. Restarting in $RestartDelaySeconds seconds..."
    Start-Sleep -Seconds $RestartDelaySeconds
  } else {
    Write-Log "Normal exit detected. Watchdog stopping."
    break
  }
}
