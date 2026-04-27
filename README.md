# PlexTraktSync Tray App

An unofficial Windows tray launcher for [PlexTraktSync](https://github.com/Taxel/PlexTraktSync). It keeps `plextraktsync watch` running in the background without a visible console window.

This was vibe-coded with Codex during a real "please just make Plex and Trakt behave again" troubleshooting session. It is not affiliated with, maintained by, or endorsed by the PlexTraktSync project; all credit goes to them for putting it together in the first place!

## What it does

- starts `plextraktsync watch`
- shows a tray icon for running, paused, and stopped state
- lets you start, stop, pause, resume, and restart the watcher
- checks whether the installed PlexTraktSync package is current, then updates it through `pipx` when needed
- opens Plex Web and Trakt directly from the tray menu
- opens the PlexTraktSync log and config folder
- restarts the watcher if it exits

## Install

### Windows Release Zip

Use this if you just want the tray app installed.

Before installing the tray app, you need:

- Windows 10 or 11
- Python 3.10 or newer
- a reachable Plex Media Server
- a Trakt account

The release zip bundles the tray app itself. Python is needed because PlexTraktSync runs through `pipx`, which is Python's recommended tool for installing command-line apps in their own isolated environments.

First install `pipx`, then install and log in to PlexTraktSync:

```powershell
py -m pip install --user pipx
py -m pipx ensurepath
pipx install PlexTraktSync
plextraktsync login
```

If you just installed `pipx` for the first time, close and reopen PowerShell before running the `pipx install PlexTraktSync` command.

Steps:

1. Go to the [latest release](https://github.com/Snorfle/plextraktsync-tray/releases/latest).
2. Download `PlexTraktSyncTray-*.zip`.
3. Extract the zip somewhere you want the app to live.
4. Open PowerShell in the extracted folder.
5. Run:

```powershell
.\install_release.ps1
```

If PowerShell blocks the script, run this instead:

```powershell
powershell -ExecutionPolicy Bypass -File .\install_release.ps1
```

The installer creates a normal user Windows logon task named `PlexTraktSync Tray` and starts the tray app. It does not run elevated.

### Developer Install

Use this if you want to edit or rebuild the app yourself.

1. Clone the repo.
2. Open PowerShell in the repo folder.
3. Run:

```powershell
.\setup_tray_app.ps1
```

## Developer Files

- `plextraktsync_tray.py` - the tray app source
- `requirements.txt` - tray app dependencies
- `setup_tray_app.ps1` - creates a venv, builds the app, and registers the logon task from source
- `install_release.ps1` - registers a packaged release build as a Windows logon task
- `build_release.ps1` - builds a zip suitable for GitHub Releases
- `LICENSE` - MIT license

## Notes

- The tray app uses your existing `pipx` install of `PlexTraktSync`.
- The update tray action checks the installed PlexTraktSync version against PyPI. When an update is available, it runs `pipx upgrade plextraktsync` and restarts the watcher.
- The Windows scheduled task created by setup is named `PlexTraktSync Tray`.
- The scheduled task launches the packaged executable, not the Python script directly.
- Do not commit your PlexTraktSync `.env`, `.pytrakt.json`, `servers.yml`, logs, cache, or packaged build output.
