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
- writes tray diagnostics to `%LOCALAPPDATA%\PlexTraktSync\PlexTraktSync\Logs\plextraktsync-tray.log`
- checks Trakt auth in the background and alerts if PlexTraktSync's saved token stops working
- falls back to marking movies watched on Trakt when Plex reports a stopped movie at 90% or later

## Experimental Branches

Simkl support is being tested separately on the
[`experimental/simkl-target`](https://github.com/Snorfle/plextraktsync-tray/tree/experimental/simkl-target)
branch. It is intentionally separate from the normal tray release and should be
treated as experimental.

## Install

### Windows Release Zip

Use this if you just want the tray app installed.

Requirements:

- Windows 10 or 11
- Python 3.10 or newer for `pipx` and PlexTraktSync
- `pipx`
- PlexTraktSync already installed and logged in with `pipx`

The release zip bundles the tray app itself, so Python is mainly needed for the underlying `pipx` install of PlexTraktSync. Developer installs have been tested with Python 3.13 and 3.14.

PlexTraktSync's recommended install path is:

```powershell
pipx install PlexTraktSync
plextraktsync login
```

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

The installer creates a normal user Windows task named `PlexTraktSync Tray`, adds a `PlexTraktSync Tray` shortcut to your Start menu, and starts the tray app. The task runs at logon and checks every minute that the tray is still running. It does not run elevated.

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
- `setup_tray_app.ps1` - creates a venv, builds the app, registers the logon task from source, and adds the Start menu shortcut
- `install_release.ps1` - registers a packaged release build as a Windows logon task and adds the Start menu shortcut
- `build_release.ps1` - builds a zip suitable for GitHub Releases
- `LICENSE` - MIT license

## Trakt Watched Fallback

PlexTraktSync normally handles Trakt scrobbling directly. The tray also watches for completed movie events and posts a watched-history fallback to Trakt if PlexTraktSync misses the final watched event. This uses PlexTraktSync's saved `.pytrakt.json` token and Plex movie IDs from `servers.yml`.

A movie counts as completed when PlexTraktSync reports `Played: True` or when playback stops at 90% or later. That second rule matches Plex's default watched threshold and covers cases where Plex has marked the movie watched but PlexTraktSync's event still says `Played: False`.

## Auth Health Checks

The tray menu shows a Trakt auth status row.

Trakt is checked by calling the Trakt API with PlexTraktSync's saved `.pytrakt.json` token. The tray refreshes OAuth credentials one day before expiry, writes the rotated credentials atomically, and restarts the watcher so it loads the new token. If refresh fails and the tray reports `Trakt auth failed`, run:

```powershell
plextraktsync trakt-login
```

Use `Check Auth Now` from the tray menu to force the check.

## Changelog

### v0.3.1

- refresh Trakt OAuth credentials one day before expiry instead of waiting for PlexTraktSync's final ten-minute window
- stop the watcher during token rotation to prevent concurrent refresh attempts from invalidating each other
- write refreshed credentials atomically and restart the watcher so it immediately loads the new token
- log OAuth refresh attempts, success, and failures in the tray diagnostics log

### v0.3.0

- added a real packaged app icon and Start menu shortcut icon
- added tray diagnostics at `%LOCALAPPDATA%\PlexTraktSync\PlexTraktSync\Logs\plextraktsync-tray.log`
- changed the Windows task to check every minute that the tray is still running
- clean up stale orphan `plextraktsync watch` processes before starting a new tray-managed watcher

### v0.2.0

- added Trakt auth health checks in the tray menu
- added a watched-history fallback for movies Plex marks watched at 90% or later
- hardened the tray supervisor so one failed status/auth/fallback check cannot kill the monitor loop
- made watched-fallback dedupe state writes atomic
- ignored malformed progress/history values instead of letting them break background sync
- tightened watcher stop/update state handling to avoid duplicate background actions
- updated the packaged release zip

### v0.1.0

- initial Windows tray app release

## Notes

- The tray app uses your existing `pipx` install of `PlexTraktSync`.
- The update tray action checks the installed PlexTraktSync version against PyPI. When an update is available, it runs `pipx upgrade plextraktsync` and restarts the watcher.
- The Windows scheduled task created by setup is named `PlexTraktSync Tray`.
- The scheduled task launches the packaged executable, not the Python script directly.
- Do not commit your PlexTraktSync `.env`, `.pytrakt.json`, `servers.yml`, logs, cache, or packaged build output.
