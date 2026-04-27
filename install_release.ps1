$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$taskName = "PlexTraktSync Tray"
$user = "$env:COMPUTERNAME\$env:USERNAME"

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
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Starts the packaged PlexTraktSync tray app at logon." -Force | Out-Null
Start-ScheduledTask -TaskName $taskName

Write-Host "PlexTraktSync Tray installed and started."
