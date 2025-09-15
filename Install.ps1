$ErrorActionPreference = 'Stop'
$AppName   = 'VismaReleaseWatcher'
$SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$MainSrc   = Join-Path $SourceDir 'VismaReleaseWatcher.ps1'
$UninstSrc = Join-Path $SourceDir 'Uninstall.ps1'
$ReadmeSrc = Join-Path $SourceDir 'README.md'
$WatchdogSrc = Join-Path $SourceDir 'Watchdog.ps1'
$WatchdogVbsSrc = Join-Path $SourceDir 'Watchdog.vbs'
$InstallDir = Join-Path $env:ProgramData $AppName

if (-not (Test-Path $MainSrc)) { throw "Main script not found: $MainSrc" }

Write-Host "Installing $AppName to $InstallDir ..."
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item $MainSrc    -Destination $InstallDir -Force
if (Test-Path $UninstSrc) { Copy-Item $UninstSrc -Destination $InstallDir -Force }
if (Test-Path $ReadmeSrc) { Copy-Item $ReadmeSrc -Destination $InstallDir -Force }
if (Test-Path $WatchdogSrc) { Copy-Item $WatchdogSrc -Destination $InstallDir -Force }
if (Test-Path $WatchdogVbsSrc) { Copy-Item $WatchdogVbsSrc -Destination $InstallDir -Force }

$TargetScript = Join-Path $InstallDir 'VismaReleaseWatcher.ps1'
$WatchdogTarget = Join-Path $InstallDir 'Watchdog.ps1'
$WatchdogVbsTarget = Join-Path $InstallDir 'Watchdog.vbs'

# Create Scheduled Task to run at user logon with interactive session (tray icon)
# Use the watchdog to supervise the tray app for resilience
$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ("`"$WatchdogVbsTarget`"")
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Compatibility Win8 -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $env:UserName -LogonType Interactive -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal

$startupShortcutCreated = $false
try {
    Register-ScheduledTask -TaskName $AppName -InputObject $task -Force | Out-Null
    Write-Host "Registered Scheduled Task '$AppName' (runs at logon)."
} catch {
    Write-Warning "Failed to register scheduled task: $($_.Exception.Message). Falling back to Startup shortcut."
    try {
        $startupDir = Join-Path $env:AppData "Microsoft\Windows\Start Menu\Programs\Startup"
        New-Item -ItemType Directory -Force -Path $startupDir | Out-Null
        $lnk = Join-Path $startupDir "$AppName.lnk"
        $ws = New-Object -ComObject WScript.Shell
        $sc = $ws.CreateShortcut($lnk)
        $sc.TargetPath = "wscript.exe"
        $sc.Arguments = "`"$WatchdogVbsTarget`""
        $sc.WorkingDirectory = $InstallDir
        $sc.IconLocation = "pwsh.exe"
        $sc.Save()
        $startupShortcutCreated = $true
        Write-Host "Created Startup shortcut at $lnk."
    } catch {
        throw "Failed to create Startup shortcut: $($_.Exception.Message)"
    }
}

# Start immediately
try {
    # Start the watchdog, which will launch and supervise the tray app
    Start-Process -FilePath 'wscript.exe' -ArgumentList ("`"$WatchdogVbsTarget`"") | Out-Null
    Write-Host "$AppName watchdog started silently (via Windows PowerShell). You should see the tray icon shortly."
} catch {
    Write-Warning "Installed, but failed to start watchdog immediately: $($_.Exception.Message)"
}

if ($startupShortcutCreated) {
    Write-Host "Autostart configured via Startup shortcut."
} else {
    Write-Host "Autostart configured via Scheduled Task."
}

Write-Host "Done."
