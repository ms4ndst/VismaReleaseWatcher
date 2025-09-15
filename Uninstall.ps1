$ErrorActionPreference = 'Stop'
$AppName = 'VismaReleaseWatcher'
$InstallDir = Join-Path $env:ProgramData $AppName

Write-Host "Removing Scheduled Task '$AppName' ..."
try {
    if (Get-ScheduledTask -TaskName $AppName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $AppName -Confirm:$false
        Write-Host "Scheduled Task removed."
    } else {
        Write-Host "Scheduled Task not found."
    }
} catch {
    Write-Warning "Could not remove Scheduled Task: $($_.Exception.Message)"
}

Write-Host "Deleting install directory $InstallDir ..."
try {
    if (Test-Path $InstallDir) {
        # Stop any running processes
        Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*VismaReleaseWatcher.ps1*' -or $_.CommandLine -like '*Watchdog.ps1*' } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
        Start-Sleep -Seconds 1
        Remove-Item -LiteralPath $InstallDir -Recurse -Force
        Write-Host "Install directory removed."
    } else {
        Write-Host "Install directory not found."
    }
} catch {
    Write-Warning "Could not remove install directory: $($_.Exception.Message)"
}

Write-Host "If the tray icon is still running, right-click it and choose Exit."