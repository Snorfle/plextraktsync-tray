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

1. Go to the experimental Simkl prerelease or build this branch locally.
2. Download `PlexTraktSyncTrayExperimental-*.zip`.
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

The installer creates a normal user Windows task named `PlexTraktSync Tray Experimental`, adds a `PlexTraktSync Tray Experimental` shortcut to your Start menu, and starts the tray app. The task runs at logon and checks every minute that the tray is still running. It does not run elevated.

This experimental build uses a separate app identity from the normal tray release. It does not replace the normal `PlexTraktSync Tray` scheduled task or Start menu shortcut.

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

## Experimental Targets

This branch includes two opt-in experimental targets in addition to the normal
PlexTraktSync/Trakt watcher:

- Serializd episode logging
- Simkl movie and episode history sync

The normal `main` release does not include these targets.

The experimental multi-target branch stores completed Plex movie and episode events in a local SQLite ledger at `%LOCALAPPDATA%\PlexTraktSyncTrayExperimental\target_sync.sqlite`.

The ledger separates media events from target attempts so targets can be handled independently. For example, a movie can be `synced` on Trakt while an episode is synced to Serializd.

This branch includes experimental target dispatchers:

- Trakt writes the existing missed-movie fallback only.
- Serializd writes completed episodes when a bearer token is available from `SERIALIZD_TOKEN` or the existing Streaming History Sync Serializd token in Windows Credential Manager.
- Simkl writes completed movies and episodes through Simkl's official API when a Simkl client ID and user token are configured.

### Experimental Serializd Target

Serializd support is opt-in and episode-only. It is based on Serializd's
current private web API behavior, not a documented public API, so treat it as
more fragile than the Simkl target.

What it syncs:

- completed Plex TV episodes
- the original Plex completion date

What it does not sync:

- movies
- ratings
- live scrobble/progress state
- deletes or reconciliation of old Serializd logs

#### Serializd Setup

The Serializd target does not include a login flow or `Connect Serializd`
button. It needs a bearer token from your own Serializd account. The simplest
setup for testers is a user environment variable:

```powershell
setx SERIALIZD_TOKEN "your-serializd-bearer-token"
```

After setting it, restart `PlexTraktSync Tray Experimental` so the scheduled
task sees the new environment variable.

If you already use Streaming History Sync on the same Windows account, this
branch can also reuse the existing Windows Credential Manager entry:

- service: `StreamingHistorySync.Serializd`
- username: `oauth`

Serializd stays inactive when no token is configured. The tray menu will show
`Serializd: not configured (...)` in that state.

### Experimental Simkl Target

Simkl support is opt-in on this experimental branch. It is not part of the normal `main` release line.

This target is meant for people who want PlexTraktSync Tray to send completed
Plex watches to Simkl. If Simkl is already connected directly to the same Plex
server through Simkl's own Plex integration, do not enable this target for the
same account unless you are intentionally testing duplicate handling.

What it syncs:

- completed Plex movies
- completed Plex TV episodes
- the original Plex completion timestamp

What it does not sync:

- ratings
- live scrobble/progress state
- rewatches as separate Simkl rewatch sessions
- deletes or reconciliation of old Simkl history

#### Simkl Setup

1. Register a Simkl application and set its client ID:

   - Open [Simkl Developer Settings](https://simkl.com/settings/developer/).
   - Create an app for this experimental tray build.
   - Use the experimental branch URL as the project or redirect URL if Simkl requires one:
     `https://github.com/Snorfle/plextraktsync-tray/tree/experimental/simkl-target`
   - Copy the app's client ID. A client secret is not needed for this PIN flow.

2. Create or edit `%LOCALAPPDATA%\PlexTraktSyncTrayExperimental\simkl_target.json`:

```powershell
New-Item -ItemType Directory -Path "$env:LOCALAPPDATA\PlexTraktSyncTrayExperimental" -Force
@{
  enabled = $false
  client_id = "your-simkl-client-id"
  account_id = $null
  username = $null
  account_type = $null
} | ConvertTo-Json | Set-Content "$env:LOCALAPPDATA\PlexTraktSyncTrayExperimental\simkl_target.json"
```

3. Start `PlexTraktSync Tray Experimental`.
4. Right-click the tray icon and choose `Connect Simkl`.
5. Approve the PIN in the browser window that opens.

The tray stores the Simkl access token in Windows Credential Manager under `PlexTraktSyncTrayExperimental.Simkl`. Non-secret Simkl settings live in `%LOCALAPPDATA%\PlexTraktSyncTrayExperimental\simkl_target.json`.

The Simkl target checks whether the exact movie or episode is already watched before writing history. It preserves the Plex completion timestamp and records the result in the local target ledger under `simkl`.

To disconnect later, use `Disconnect Simkl` from the experimental tray menu. That removes the stored Simkl token and disables the Simkl target without touching PlexTraktSync's Trakt login.

### Experimental Build Identity

The Simkl branch intentionally uses separate Windows identity anchors:

- app name: `PlexTraktSync Tray Experimental`
- scheduled task: `PlexTraktSync Tray Experimental`
- Start menu shortcut: `PlexTraktSync Tray Experimental`
- packaged app folder/exe: `PlexTraktSyncTrayExperimental`
- app mutex: `PlexTraktSyncTrayExperimentalApp`
- Simkl Credential Manager service: `PlexTraktSyncTrayExperimental.Simkl`
- tray diagnostics, target ledger, and Simkl settings: `%LOCALAPPDATA%\PlexTraktSyncTrayExperimental`

It still reads PlexTraktSync's existing watch log and config from `%LOCALAPPDATA%\PlexTraktSync\PlexTraktSync`, because those belong to the underlying PlexTraktSync install.

## Auth Health Checks

The tray menu shows a Trakt auth status row.

Trakt is checked by calling the Trakt API with PlexTraktSync's saved `.pytrakt.json` token. If it reports `Trakt auth failed`, run:

```powershell
plextraktsync trakt-login
```

Use `Check Auth Now` from the tray menu to force the check.

## Changelog

### Experimental Simkl branch

- added an opt-in Simkl target using Simkl's official API
- added Simkl PIN connection, Credential Manager token storage, and disconnect support
- records Simkl target outcomes in the multi-target ledger
- split the experimental Windows app identity from the normal tray release
- accepts Simkl history responses that confirm completion in `added.statuses`
- summarizes only the latest target result for each item

### v0.3.0

- added a real packaged app icon and Start menu shortcut icon
- added tray diagnostics at `%LOCALAPPDATA%\PlexTraktSync\PlexTraktSync\Logs\plextraktsync-tray.log`
- changed the Windows task to check every minute that the tray is still running
- clean up stale orphan `plextraktsync watch` processes before starting a new tray-managed watcher
- ignored local Trakt and Serializd CSV exports

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
- The Windows scheduled task created by setup is named `PlexTraktSync Tray Experimental` on this branch.
- The scheduled task launches the packaged executable, not the Python script directly.
- Do not commit your PlexTraktSync `.env`, `.pytrakt.json`, `servers.yml`, logs, cache, or packaged build output.
