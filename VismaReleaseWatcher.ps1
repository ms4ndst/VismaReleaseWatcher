param(
    [switch]$DebugMode,
    [switch]$CheckNow,
    [switch]$ShowBalloon
)

# Visma Release Watcher - Background tray app
# - Checks https://releasenotes.control.visma.com/ twice a day using the WordPress REST API
# - Uses ETag (If-None-Match) to avoid unnecessary downloads; detects both new and edited posts
# - Tray icon: green = no updates, yellow = updates detected since last check
# - Settings GUI to set two daily check times (HH:mm, 24-hour)
# - Logs checks to CSV under ProgramData
# - Config persists LastSignature, LastLinks, LastChecked, and ApiETag

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[AppDomain]::CurrentDomain.add_UnhandledException({ param($s,$e) try { Write-Log ("UnhandledException: {0}" -f $e.ExceptionObject) } catch {} ; $script:ExitCode = 1 })
[System.Windows.Forms.Application]::add_ThreadException({ param($s,$e) try { Write-Log ("ThreadException: {0}: {1}`n{2}" -f $e.Exception.GetType().FullName, $e.Exception.Message, $e.Exception.StackTrace) } catch {} ; $script:ExitCode = 1 })

$ErrorActionPreference = 'Stop'

# Ensure modern TLS
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$AppName    = 'VismaReleaseWatcher'
$DataDir    = Join-Path $env:ProgramData $AppName
$ConfigPath = Join-Path $DataDir 'config.json'
$CsvPath    = Join-Path $DataDir 'updates.csv'
$LogPath    = Join-Path $DataDir 'app.log'

# Single instance guard
$createdNew = $false
# Use a single backslash for the Local\ namespace in .NET named mutex
$mutex = New-Object System.Threading.Mutex($true, 'Local\VismaReleaseWatcher_Mutex', [ref]$createdNew)
if (-not $createdNew) { Write-Log "Existing instance detected. Exiting."; exit 0 }

# Ensure data directory
if (-not (Test-Path $DataDir)) {
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
}

function Write-Log([string]$msg) {
    try {
        $ts = (Get-Date).ToString('s')
        "$ts $msg" | Add-Content -Path $LogPath -Encoding UTF8
    } catch {}
}

function Save-Config($config) {
    ($config | ConvertTo-Json -Depth 5) | Set-Content -Encoding UTF8 -Path $ConfigPath
}

function Load-Config() {
    if (Test-Path $ConfigPath) {
        try {
            $cfg = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
            if (-not ($cfg.PSObject.Properties.Name -contains 'LastLinks')) {
                $cfg | Add-Member -NotePropertyName LastLinks -NotePropertyValue @()
            }
            if (-not ($cfg.PSObject.Properties.Name -contains 'ApiETag')) {
                $cfg | Add-Member -NotePropertyName ApiETag -NotePropertyValue ''
            }
            return $cfg
        } catch {}
    }
    # Default config
    return [pscustomobject]@{
        CheckTimes    = @('09:00','15:00')  # 24h format HH:mm
        LastSignature = ''
        LastChecked   = ''
        LastLinks     = @()
        ApiETag       = ''
    }
}

$config = Load-Config
Write-Log "Startup. Args: CheckNow=$CheckNow ShowBalloon=$ShowBalloon"
if (-not $config.CheckTimes -or $config.CheckTimes.Count -lt 2) {
    $config.CheckTimes = @('09:00','15:00')
    Save-Config $config
}

function Initialize-Csv {
    if (-not (Test-Path $CsvPath)) {
        'Timestamp,Status,NewLinksCount,Signature' | Set-Content -Encoding UTF8 -Path $CsvPath
    }
}
Initialize-Csv

# Draw simple colored circle icons dynamically so we don't need external .ico files
function New-ColorIcon([System.Drawing.Color]$Color){
    $bmp = New-Object System.Drawing.Bitmap 16,16
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush $Color
    $g.FillEllipse($brush,1,1,14,14)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Black)
    $g.DrawEllipse($pen,1,1,14,14)
    $hicon = $bmp.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hicon)
    $g.Dispose(); $brush.Dispose(); $pen.Dispose(); $bmp.Dispose()
    return $icon
}

$greenIcon  = New-ColorIcon([System.Drawing.Color]::FromArgb(0,180,0))
$yellowIcon = New-ColorIcon([System.Drawing.Color]::FromArgb(255,200,0))

$script:ExitCode = 0
# Notify icon
$script:notify = $null
if (-not $CheckNow) {
    $script:notify = New-Object System.Windows.Forms.NotifyIcon
    $script:notify.Visible = $true
    $script:notify.Text = $AppName
    $script:notify.Icon = $greenIcon

    # Context menu
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
$checkNowItem = $menu.Items.Add('Check Now')
$settingsItem = $menu.Items.Add('Settings...')
$openDataItem = $menu.Items.Add('Open Data Folder')
$openSiteItem = $menu.Items.Add('Open Release Notes')
$showLogsItem = $menu.Items.Add('Show Logs')
$restartItem = $menu.Items.Add('Restart')
$exitItem = $menu.Items.Add('Exit')
$script:notify.ContextMenuStrip = $menu

    if ($ShowBalloon) {
        try {
            $next = (Get-NextDue).ToString('t')
            $script:notify.BalloonTipTitle = $AppName
            $script:notify.BalloonTipText = "Started. Next check: $next"
            $script:notify.ShowBalloonTip(5000)
        } catch {}
    }
}

# Simple helpers
function Get-Links([string]$html) {
    $matches = [regex]::Matches($html, 'href="(https?://releasenotes\.control\.visma\.com[^"#?]*?)"', 'IgnoreCase')
    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach($m in $matches){ [void]$set.Add($m.Groups[1].Value) }
    return ,$set
}

function Get-Signature([string]$text){
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }
    -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function Get-LinksSignature($linksSet){
    # Build a stable signature from the set of links only (avoids dynamic HTML noise)
    $arr = @($linksSet)
    $sorted = $arr | Sort-Object -Unique
    $joined = [string]::Join("`n", $sorted)
    return Get-Signature $joined
}

# Check-ForUpdates
# - Queries WP REST: /wp-json/wp/v2/posts?per_page=25&_fields=id,link,modified_gmt
# - Sends If-None-Match with persisted ApiETag; if 304 -> no change
# - Computes signature from sorted lines of "id|modified_gmt|link" to capture edits as well
# - Only notifies as UpdateDetected when new links appear vs. the previous snapshot (consistent with prior behavior)
function Check-ForUpdates {
    param([switch]$Interactive)
    try {
        $restUrl = 'https://releasenotes.control.visma.com/wp-json/wp/v2/posts?per_page=25&_fields=id,link,modified_gmt'
        $headers = @{ 'User-Agent' = 'VismaReleaseWatcher/1.1 (+PowerShell)' }
        if ($config.PSObject.Properties.Name -notcontains 'ApiETag') {
            $config | Add-Member -NotePropertyName ApiETag -NotePropertyValue ''
        }
        if ($config.ApiETag) { $headers['If-None-Match'] = $config.ApiETag }

        $resp = $null
        $status = 200
        try {
            $resp = Invoke-WebRequest -Uri $restUrl -TimeoutSec 30 -Headers $headers -ErrorAction Stop
            $status = $resp.StatusCode
        } catch {
            $we = $_.Exception
            if ($we -and $we.Response -and ($we.Response.StatusCode.value__ -eq 304)) {
                $status = 304
            } else {
                throw
            }
        }

        $sig = $config.LastSignature
        $links = $config.LastLinks
        $newCount = 0
        $hasUpdate = $false

        if ($status -eq 304) {
            # No change on server since previous ETag
            $hasUpdate = $false
        } else {
            # Update ETag if present
            try {
                $etag = $resp.Headers.ETag
                if ($etag) { $config.ApiETag = $etag }
            } catch {}

            # Parse JSON and compute signature from id|modified_gmt|link
            $json = $resp.Content | ConvertFrom-Json
            $rows = @()
            $currLinks = @()
            foreach ($p in $json) {
                $rows += ('{0}|{1}|{2}' -f $p.id, $p.modified_gmt, $p.link)
                $currLinks += $p.link
            }
            $sorted = $rows | Sort-Object
            $joined = [string]::Join("`n", $sorted)
            $sig = Get-Signature $joined

            $prevSig = $config.LastSignature
            $prevLinks = @()
            if ($config.PSObject.Properties.Name -contains 'LastLinks') { $prevLinks = @($config.LastLinks) }

            if ($prevSig) {
                if ($sig -ne $prevSig) {
                    # Compute actually new posts compared to last snapshot (by link)
                    $newLinks = @()
                    foreach ($l in $currLinks) { if (-not ($prevLinks -contains $l)) { $newLinks += $l } }
                    $newCount = $newLinks.Count
                    $hasUpdate = ($newCount -gt 0)
                }
            } else {
                # First run establishes baseline only (no update)
                $hasUpdate = $false
            }

            # Persist current snapshot
            $links = $currLinks
        }

        $config.LastSignature = $sig
        if (-not ($config.PSObject.Properties.Name -contains 'LastLinks')) {
            $config | Add-Member -NotePropertyName LastLinks -NotePropertyValue @()
        }
        $config.LastLinks = @($links)
        $config.LastChecked = (Get-Date).ToString('o')
        Save-Config $config

        $statusText = if ($hasUpdate) { 'UpdateDetected' } else { 'NoChange' }
        "$([DateTime]::Now.ToString('s')),$statusText,$newCount,$sig" | Add-Content -Encoding UTF8 -Path $CsvPath
        Write-Log "Check completed. Status=$statusText NextDue=$((Get-NextDue).ToString('s'))"

        if ($hasUpdate) {
            if ($script:notify) {
                $script:notify.Icon = $yellowIcon
                $script:notify.BalloonTipTitle = 'Visma Release Notes'
                $script:notify.BalloonTipText = 'New updates detected since last check.'
                $script:notify.ShowBalloonTip(5000)
            }
        } else {
            if ($script:notify) {
                $script:notify.Icon = $greenIcon
                if ($Interactive) {
                    $script:notify.BalloonTipTitle = 'Visma Release Notes'
                    $script:notify.BalloonTipText = 'No new updates.'
                    $script:notify.ShowBalloonTip(3000)
                }
            }
        }
    } catch {
        if ($script:notify) {
            $script:notify.BalloonTipTitle = 'Visma Release Watcher'
            $script:notify.BalloonTipText = "Error checking updates: $($_.Exception.Message)"
            $script:notify.ShowBalloonTip(5000)
        } else {
            Write-Host "Error checking updates: $($_.Exception.Message)"
        }
    }
}

function Parse-TimeStr([string]$t) {
    # returns DateTime for today at time string HH:mm
    [DateTime]::Today.Add([TimeSpan]::Parse($t))
}

function Get-NextDue {
    $now = Get-Date
    $times = @()
    foreach($t in $config.CheckTimes) {
        try {
            $dt = Parse-TimeStr $t
            if ($dt -le $now) { $dt = $dt.AddDays(1) }
            $times += $dt
        } catch {}
    }
    if ($times.Count -eq 0) { return $now.AddHours(12) }
    $times | Sort-Object | Select-Object -First 1
}

$script:timer = $null
$script:winTimer = $null
function Test-NotifyIcon {
    if (-not $script:notify) { return }
    try {
        # Probing properties can throw if icon handle becomes invalid
        $null = $script:notify.Visible
        if (-not $script:notify.Visible) { $script:notify.Visible = $true }
    } catch {
        Write-Log "NotifyIcon invalid; recreating"
        try { $script:notify.Dispose() } catch {}
        $script:notify = New-Object System.Windows.Forms.NotifyIcon
        $script:notify.Visible = $true
        $script:notify.Text = $AppName
        $script:notify.Icon = $greenIcon
        # Rebuild menu
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        $checkNowItem = $menu.Items.Add('Check Now')
        $settingsItem = $menu.Items.Add('Settings...')
        $openDataItem = $menu.Items.Add('Open Data Folder')
        $openSiteItem = $menu.Items.Add('Open Release Notes')
        $showLogsItem = $menu.Items.Add('Show Logs')
        $restartItem = $menu.Items.Add('Restart')
        $exitItem = $menu.Items.Add('Exit')
        $script:notify.ContextMenuStrip = $menu
        # Rehook clicks
        $checkNowItem.add_Click({ Check-ForUpdates -Interactive })
        $settingsItem.add_Click({ Show-Settings })
        $openDataItem.add_Click({ Start-Process explorer.exe $DataDir })
        $openSiteItem.add_Click({ Start-Process "https://releasenotes.control.visma.com/" })
        $showLogsItem.add_Click({ if (Test-Path $LogPath) { Start-Process notepad.exe $LogPath } else { [System.Windows.Forms.MessageBox]::Show("No log file yet.", $AppName) } })
        $restartItem.add_Click({ $script:ExitCode = 1; if ($script:notify) { $script:notify.Visible = $false; $script:notify.Dispose() }; if ($script:winTimer) { $script:winTimer.Stop(); $script:winTimer.Dispose() }; try { $mutex.ReleaseMutex() } catch {}; try { if ($script:hostForm) { $script:hostForm.Close() } else { [System.Windows.Forms.Application]::Exit() } } catch { [System.Windows.Forms.Application]::Exit() } })
        $exitItem.add_Click({ $script:ExitCode = 200; if ($script:notify) { $script:notify.Visible = $false; $script:notify.Dispose() }; if ($script:winTimer) { $script:winTimer.Stop(); $script:winTimer.Dispose() }; try { $mutex.ReleaseMutex() } catch {}; try { if ($script:hostForm) { $script:hostForm.Close() } else { [System.Windows.Forms.Application]::Exit() } } catch { [System.Windows.Forms.Application]::Exit() } })
    }
}

function Schedule-Next {
    if ($script:timer) { $script:timer.Dispose(); $script:timer = $null }
    if ($script:winTimer) { $script:winTimer.Stop(); $script:winTimer.Dispose(); $script:winTimer = $null }
    $due = Get-NextDue
    $now = Get-Date
    $script:schedMs = [int][Math]::Max(1000, ($due - $now).TotalMilliseconds)

    # Use a WinForms timer to fire on the UI thread to avoid cross-thread exceptions
    $script:winTimer = New-Object System.Windows.Forms.Timer
    $script:winTimer.Interval = [Math]::Min($script:schedMs, [int][uint16]::MaxValue)  # cap to ~65s; reschedule if needed

    $script:schedElapsed = 0
    $tick = {
        $script:schedElapsed += $script:winTimer.Interval
        if ($script:schedElapsed -ge $script:schedMs) {
            $script:winTimer.Stop()
            try { Check-ForUpdates } finally { Schedule-Next }
        }
    }
    $script:winTimer.add_Tick($tick)
    $script:winTimer.Start()

    Test-NotifyIcon
    if ($script:notify) { $script:notify.Text = "$AppName - Next check: $($due.ToString('t'))" }
    Write-Log "Scheduled next check at $($due.ToString('s')) (in $([int]($due - $now).TotalMinutes) min)"
}

function Show-Settings {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$AppName Settings"
    $form.Size = New-Object System.Drawing.Size(320,220)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label1 = New-Object System.Windows.Forms.Label
    $label1.Text = 'Check time 1:'
    $label1.Location = New-Object System.Drawing.Point(20,20)
    $label1.AutoSize = $true
    $form.Controls.Add($label1)

    $dtp1 = New-Object System.Windows.Forms.DateTimePicker
    $dtp1.Format = 'Time'
    $dtp1.ShowUpDown = $true
    $dtp1.Location = New-Object System.Drawing.Point(140,16)
    $dtp1.Width = 120
    $form.Controls.Add($dtp1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Text = 'Check time 2:'
    $label2.Location = New-Object System.Drawing.Point(20,60)
    $label2.AutoSize = $true
    $form.Controls.Add($label2)

    $dtp2 = New-Object System.Windows.Forms.DateTimePicker
    $dtp2.Format = 'Time'
    $dtp2.ShowUpDown = $true
    $dtp2.Location = New-Object System.Drawing.Point(140,56)
    $dtp2.Width = 120
    $form.Controls.Add($dtp2)

    try {
        if ($config.CheckTimes.Count -ge 2) {
            $dtp1.Value = Parse-TimeStr ($config.CheckTimes[0])
            $dtp2.Value = Parse-TimeStr ($config.CheckTimes[1])
        }
    } catch {}

    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = 'Save'
    $saveBtn.Location = New-Object System.Drawing.Point(60,110)
    $saveBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($saveBtn)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = 'Cancel'
    $cancelBtn.Location = New-Object System.Drawing.Point(160,110)
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelBtn)

    $form.AcceptButton = $saveBtn
    $form.CancelButton = $cancelBtn

    if ($form.ShowDialog() -eq 'OK') {
        $config.CheckTimes = @(
            $dtp1.Value.ToString('HH:mm'),
            $dtp2.Value.ToString('HH:mm')
        )
        Save-Config $config
        Schedule-Next
        if ($script:notify) {
            $script:notify.BalloonTipTitle = $AppName
            $script:notify.BalloonTipText = "Schedule updated. Next: $((Get-NextDue).ToString('g'))"
            $script:notify.ShowBalloonTip(3000)
        }
    }
}

if (-not $CheckNow) {
    # Hook up menu events
    $checkNowItem.add_Click({ Check-ForUpdates -Interactive })
    $settingsItem.add_Click({ Show-Settings })
    $openDataItem.add_Click({ Start-Process explorer.exe $DataDir })
    $openSiteItem.add_Click({ Start-Process "https://releasenotes.control.visma.com/" })
    $showLogsItem.add_Click({ if (Test-Path $LogPath) { Start-Process notepad.exe $LogPath } else { [System.Windows.Forms.MessageBox]::Show("No log file yet.", $AppName) } })
    $restartItem.add_Click({ $script:ExitCode = 1; if ($script:notify) { $script:notify.Visible = $false; $script:notify.Dispose() }; if ($script:winTimer) { $script:winTimer.Stop(); $script:winTimer.Dispose() }; try { $mutex.ReleaseMutex() } catch {}; try { if ($script:hostForm) { $script:hostForm.Close() } else { [System.Windows.Forms.Application]::Exit() } } catch { [System.Windows.Forms.Application]::Exit() } })
    $exitItem.add_Click({
        $script:ExitCode = 200  # signal to watchdog that this was an intentional exit
        if ($script:notify) { $script:notify.Visible = $false; $script:notify.Dispose() }
        if ($script:timer) { $script:timer.Dispose() }
        try { $mutex.ReleaseMutex() } catch {}
        try { if ($script:hostForm) { $script:hostForm.Close() } else { [System.Windows.Forms.Application]::Exit() } } catch { [System.Windows.Forms.Application]::Exit() }
    })

    # Double-click to open Settings
    $script:notify.add_DoubleClick({ Show-Settings })
    $script:notify.add_BalloonTipClicked({ Write-Log 'Balloon clicked'; Show-Settings })
    $script:notify.add_BalloonTipClosed({ Write-Log 'Balloon closed' })

    # Start scheduling (no immediate check to avoid out-of-schedule notifications)
    Schedule-Next

# Create an invisible host form to keep the message loop alive reliably
$script:hostForm = New-Object System.Windows.Forms.Form
$script:hostForm.Text = $AppName
$script:hostForm.ShowInTaskbar = $false
$script:hostForm.FormBorderStyle = 'FixedToolWindow'
$script:hostForm.Opacity = 0
$script:hostForm.WindowState = 'Minimized'
$script:hostForm.Size = [System.Drawing.Size]::new(0,0)
$script:hostForm.StartPosition = 'Manual'

try {
    Write-Log "Entering message loop (host form)"
    [System.Windows.Forms.Application]::Run($script:hostForm)
    Write-Log "Message loop returned"
} catch {
    Write-Log "Unhandled exception in message loop: $($_.Exception.GetType().FullName): $($_.Exception.Message)"
    $script:ExitCode = 1
} finally {
    Write-Log "Application loop exited. ExitCode=$script:ExitCode"
    try { if ($script:notify) { $script:notify.Visible = $false; $script:notify.Dispose() } } catch {}
    try { if ($script:winTimer) { $script:winTimer.Stop(); $script:winTimer.Dispose() } } catch {}
    try { $mutex.ReleaseMutex() } catch {}
    [Environment]::Exit($script:ExitCode)
}
} else {
    # Headless one-shot check
    Check-ForUpdates
    try { $mutex.ReleaseMutex() } catch {}
    exit 0
}
