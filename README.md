# Visma Release Watcher

A small PowerShell tray application that checks https://releasenotes.control.visma.com/ twice a day and indicates status via a system tray icon:
- Green = no change since last check
- Yellow = new updates detected since last check

Checks and results are logged to CSV and settings are persisted in ProgramData.

## Install

1) Ensure PowerShell can run local scripts (per-session):

   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

2) From this folder, run:

   pwsh -NoProfile -ExecutionPolicy Bypass -File ./Install.ps1

This installs files to:

- C:\\ProgramData\\VismaReleaseWatcher\\VismaReleaseWatcher.ps1 (tray app)
- C:\\ProgramData\\VismaReleaseWatcher\\Watchdog.ps1 (supervisor)
- C:\\ProgramData\\VismaReleaseWatcher\\Watchdog.vbs (silent launcher)

Autostart behavior:
- The installer first attempts to register a per-user Scheduled Task. If access is denied, it falls back to a Startup shortcut that launches wscript.exe with Watchdog.vbs (fully hidden).
- The watchdog launches the tray app using Windows PowerShell (powershell.exe) in hidden mode.

The app starts immediately after install (hidden) and the tray icon appears (check the hidden-icons overflow if not pinned).

## Uninstall

Run:

   pwsh -NoProfile -ExecutionPolicy Bypass -File ./Uninstall.ps1

Then exit the tray app if still running (right-click tray icon > Exit). If the tray icon is gone but autostart remains, delete the Startup shortcut:

- C:\\Users\\<you>\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\VismaReleaseWatcher.lnk

## Usage

Right-click the tray icon for menu options:
- Check Now: immediately checks for updates
- Settings: GUI to set two daily check times (HH:mm, 24-hour)
- Open Data Folder: opens ProgramData\\VismaReleaseWatcher
- Open Release Notes: opens the release notes website
- Show Logs: opens app.log in Notepad
- Restart: intentionally restarts via the watchdog
- Exit: quits the app (won’t auto-restart)

First run establishes a baseline signature and will typically show green until content changes.

## Data and Logs

- Config: C:\\ProgramData\\VismaReleaseWatcher\\config.json
- CSV log: C:\\ProgramData\\VismaReleaseWatcher\\updates.csv
- App log: C:\\ProgramData\\VismaReleaseWatcher\\app.log
- Watchdog log: C:\\ProgramData\\VismaReleaseWatcher\\watchdog.log

CSV columns: Timestamp, Status (NoChange|UpdateDetected), NewLinksCount, Signature

## Troubleshooting

- Tray icon not visible: click the taskbar up-arrow to show hidden icons and drag the icon to pin it.
- Icon disappeared: it should be recreated automatically; if not, right-click > Restart.
- A console window appears: the launcher uses wscript + powershell.exe hidden. If you still see a console, ensure no duplicate Startup entries exist that run pwsh.exe directly.
- Manual start: run the tray app once in the foreground for diagnostics:

  pwsh -NoProfile -STA -ExecutionPolicy Bypass -File "C:\\ProgramData\\VismaReleaseWatcher\\VismaReleaseWatcher.ps1" -ShowBalloon

## Notes

- Requires Windows with .NET Windows Forms available (PowerShell 7+ or Windows PowerShell is fine; launcher uses Windows PowerShell for hidden startup).
- The app dynamically draws the tray icons; no .ico files are required.
- Network errors are shown as a tray balloon but do not stop the app.
