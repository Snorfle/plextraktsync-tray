$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvDir = Join-Path $baseDir ".venv"
$pythonExe = Join-Path $venvDir "Scripts\python.exe"
$requirements = Join-Path $baseDir "requirements.txt"
$appScript = Join-Path $baseDir "plextraktsync_tray.py"
$appExe = Join-Path $baseDir "dist\PlexTraktSyncTray\PlexTraktSyncTray.exe"
$taskName = "PlexTraktSync Tray"
$user = "$env:COMPUTERNAME\$env:USERNAME"

if (-not (Test-Path $venvDir)) {
    py -3 -m venv $venvDir
}

& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r $requirements
& $pythonExe -m pip install pyinstaller
& $pythonExe -m PyInstaller --noconfirm --windowed --name PlexTraktSyncTray $appScript

try {
    Unregister-ScheduledTask -TaskName "PlexTraktSync Watch" -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

$action = New-ScheduledTaskAction -Execute $appExe
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Starts the PlexTraktSync tray app at logon." -Force | Out-Null
Start-ScheduledTask -TaskName $taskName

Write-Host "Tray app installed and started."
