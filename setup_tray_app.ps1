$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvDir = Join-Path $baseDir ".venv"
$pythonExe = Join-Path $venvDir "Scripts\python.exe"
$requirements = Join-Path $baseDir "requirements.txt"
$appScript = Join-Path $baseDir "plextraktsync_tray.py"
$appIcon = Join-Path $baseDir "assets\PlexTraktSyncTray.ico"
$appExe = Join-Path $baseDir "dist\PlexTraktSyncTray\PlexTraktSyncTray.exe"
$taskName = "PlexTraktSync Tray"
$shortcutName = "PlexTraktSync Tray.lnk"
$programsDir = [Environment]::GetFolderPath("Programs")
$shortcutPath = Join-Path $programsDir $shortcutName
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
    $shortcut.IconLocation = "$appIcon,0"
    $shortcut.Description = "Launch PlexTraktSync Tray"
    $shortcut.Save()
}

if (-not (Test-Path $venvDir)) {
    py -3 -m venv $venvDir
}

& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r $requirements
& $pythonExe -m pip install pyinstaller
& $pythonExe -m PyInstaller --noconfirm --windowed --name PlexTraktSyncTray --icon $appIcon $appScript

# Clean up the old pre-tray task name if it exists.
try {
    Unregister-ScheduledTask -TaskName "PlexTraktSync Watch" -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

$action = New-ScheduledTaskAction -Execute $appExe
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$watchdogTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 0) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($logonTrigger, $watchdogTrigger) -Principal $principal -Settings $settings -Description "Starts the PlexTraktSync tray app at logon and checks every minute that it is still running." -Force | Out-Null
Install-StartMenuShortcut -TargetPath $appExe
Start-ScheduledTask -TaskName $taskName

Write-Host "Tray app installed, added to the Start menu, and started."
