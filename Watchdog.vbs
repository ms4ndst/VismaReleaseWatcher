Set shell = CreateObject("Wscript.Shell")
cmd = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\ProgramData\VismaReleaseWatcher\Watchdog.ps1"" -RestartDelaySeconds 10"
shell.Run cmd, 0, False
