$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = "PlexTraktSync Tray"
$shortcutName = "PlexTraktSync Tray.lnk"
$programsDir = [Environment]::GetFolderPath("Programs")
$shortcutPath = Join-Path $programsDir $shortcutName
$appIcon = Join-Path $baseDir "assets\PlexTraktSyncTray.ico"
$user = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { "$env:COMPUTERNAME\$env:USERNAME" }

function Install-StartMenuShortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    if (-not (Test-Path $programsDir)) {
        New-Item -ItemType Directory -Path $programsDir -Force | Out-Null
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = Split-Path -Parent $TargetPath
    if (Test-Path $appIcon) {
        $shortcut.IconLocation = "$appIcon,0"
    } else {
        $shortcut.IconLocation = "$TargetPath,0"
    }
    $shortcut.Description = "Launch PlexTraktSync Tray"
    $shortcut.Save()
}

$exeCandidates = @(
    (Join-Path $baseDir "PlexTraktSyncTray\PlexTraktSyncTray.exe"),
    (Join-Path $baseDir "dist\PlexTraktSyncTray\PlexTraktSyncTray.exe"),
    (Join-Path $baseDir "PlexTraktSyncTray.exe")
)

$appExe = $exeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $appExe) {
    throw "Could not find PlexTraktSyncTray.exe next to this installer."
}

$watcherPython = Join-Path $env:USERPROFILE "pipx\venvs\plextraktsync\Scripts\python.exe"
if (-not (Test-Path $watcherPython)) {
    throw "PlexTraktSync was not found at $watcherPython. Install and configure PlexTraktSync with pipx first, then run this installer again."
}

try {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
} catch {}

try {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

$action = New-ScheduledTaskAction -Execute $appExe
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$watchdogTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 0) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($logonTrigger, $watchdogTrigger) -Principal $principal -Settings $settings -Description "Starts the packaged PlexTraktSync tray app at logon and checks every minute that it is still running." -Force | Out-Null
Install-StartMenuShortcut -TargetPath $appExe
Start-ScheduledTask -TaskName $taskName

Write-Host "PlexTraktSync Tray installed, added to the Start menu, and started."
