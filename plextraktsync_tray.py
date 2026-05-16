import ctypes
import json
import os
import re
import signal
import shutil
import subprocess
import threading
import time
import urllib.parse
import urllib.error
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pystray
from PIL import Image, ImageDraw
from pystray import MenuItem as Item


APP_NAME = "PlexTraktSync Tray"
TASK_NAME = "PlexTraktSync Tray"
BASE_DIR = Path(__file__).resolve().parent
LOCAL_APPDATA = Path(os.environ["LOCALAPPDATA"]) / "PlexTraktSync" / "PlexTraktSync"
PLEXTRAKTSYNC_PYTHON = Path.home() / "pipx" / "venvs" / "plextraktsync" / "Scripts" / "python.exe"
LOG_FILE = LOCAL_APPDATA / "Logs" / "plextraktsync.log"
PYTRAKT_CONFIG_FILE = LOCAL_APPDATA / ".pytrakt.json"
SERVERS_CONFIG_FILE = LOCAL_APPDATA / "servers.yml"
COMPLETED_MOVIE_STATE_FILE = LOCAL_APPDATA / "completed_movie_sync_state.json"
CHECK_INTERVAL_SECONDS = 10
RESTART_DELAY_SECONDS = 15
LOG_TAIL_BYTES = 65536
CREATE_NO_WINDOW = 0x08000000
ERROR_ALREADY_EXISTS = 183
MUTEX_NAME = "Local\\PlexTraktSyncTrayApp"
PLEX_BASE_URL = "http://127.0.0.1:32400"
PLEX_WEB_URL = f"{PLEX_BASE_URL}/web"
TRAKT_WEB_URL = "https://trakt.tv/"
TRAKT_API_SETTINGS_URL = "https://api.trakt.tv/users/settings"
PYPI_JSON_URL = "https://pypi.org/pypi/plextraktsync/json"
PLAYBACK_STALE_MINUTES = 30
WATCHED_PROGRESS_THRESHOLD = 90.0
TRAKT_AUTH_CHECK_SECONDS = 15 * 60
TRAKT_AUTH_RETRY_SECONDS = 60
ON_PLAY_PATTERN = re.compile(
    r"^(?P<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+\s+INFO\[.*?\]:on_play: <[^>]+:(?P<title>.+)>: (?P<progress>\d+(?:\.\d+)?)%, State: (?P<state>\w+)",
)
WATCHED_MOVIE_PATTERN = re.compile(
    r"^(?P<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+\s+INFO\[.*?\]:on_play: "
    r"<(?P<kind>[^:>]+):(?P<rating_key>\d+):(?P<title>.+)>: (?P<progress>\d+(?:\.\d+)?)%, "
    r"State: (?P<state>\w+), Played: (?P<played>True|False)",
)

@dataclass(frozen=True)
class WatchedMovieEvent:
    timestamp: datetime
    rating_key: str
    title: str
    progress: float


class CompletedMovieSync:
    """Fallback watched-history sync for completed Plex movie watch events."""

    def __init__(self, log_path: Path, state_path: Path) -> None:
        self.log_path = log_path
        self.state_path = state_path
        self.lock = threading.Lock()
        self.running = False
        self.last_status = "Trakt fallback: waiting"
        self.synced_keys = self._load_synced_keys()

    def status_text(self) -> str:
        with self.lock:
            running = self.running
            last_status = self.last_status
        if running:
            return "Trakt fallback: syncing"
        return last_status

    def check_log(self) -> None:
        with self.lock:
            if self.running:
                return
        event = self._latest_unsynced_event()
        if event is None:
            self.last_status = "Trakt fallback: waiting"
            return

        thread = threading.Thread(target=self._sync_event, args=(event,), daemon=True)
        thread.start()

    def _latest_unsynced_event(self) -> WatchedMovieEvent | None:
        if not self.log_path.exists():
            return None

        try:
            with self.log_path.open("rb") as log_file:
                log_file.seek(0, os.SEEK_END)
                size = log_file.tell()
                log_file.seek(max(0, size - LOG_TAIL_BYTES))
                lines = log_file.read().decode("utf-8", errors="replace").splitlines()
        except OSError:
            self.last_status = "Trakt fallback: log unavailable"
            return None

        for line in reversed(lines[-500:]):
            event = completed_movie_event_from_log_line(line)
            if event is None:
                continue
            sync_key = self._sync_key(event.rating_key, event.timestamp)
            if sync_key not in self.synced_keys:
                return event
        return None

    def _sync_event(self, event: WatchedMovieEvent) -> None:
        with self.lock:
            if self.running:
                return
            self.running = True

        try:
            self.last_status = f"Trakt fallback: logging {event.title}"
            mark_trakt_movie_watched(event)
            sync_key = self._sync_key(event.rating_key, event.timestamp)
            self.synced_keys.add(sync_key)
            try:
                self._save_synced_keys()
            except OSError as exc:
                self.last_status = f"Trakt fallback: logged {event.title}; state save failed"
                notify_message(f"Marked {event.title} watched on Trakt, but local dedupe state was not saved: {friendly_error(exc)}.")
                return
            self.last_status = f"Trakt fallback: logged {event.title}"
            notify_message(f"Marked {event.title} watched on Trakt.")
        except Exception as exc:
            self.last_status = f"Trakt fallback failed: {friendly_error(exc)}"
            notify_message(f"Trakt watched fallback failed: {friendly_error(exc)}.")
        finally:
            with self.lock:
                self.running = False
            refresh_icon()

    def _load_synced_keys(self) -> set[str]:
        if not self.state_path.exists():
            return set()
        try:
            payload = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return set()
        return {str(item) for item in payload.get("synced", [])}

    def _save_synced_keys(self) -> None:
        cutoff = (datetime.now() - timedelta(days=180)).date().isoformat()
        pruned = sorted(
            key
            for key in self.synced_keys
            if ":" in key and key.split(":", 1)[1][:10] >= cutoff
        )
        self.synced_keys = set(pruned)
        payload = {"synced": pruned}
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = self.state_path.with_name(f"{self.state_path.name}.tmp")
        temp_path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
        temp_path.replace(self.state_path)

    @staticmethod
    def _sync_key(rating_key: str, timestamp: datetime) -> str:
        return f"{rating_key}:{timestamp.isoformat(timespec='minutes')}"


class WatcherManager:
    """Owns the `plextraktsync watch` child process and tray-facing state."""

    def __init__(self) -> None:
        self.process: subprocess.Popen | None = None
        self.auto_restart = True
        self.lock = threading.Lock()
        self.last_error: str | None = None
        self.last_start_time = 0.0
        self.last_connected_at: float | None = None
        self.last_update_result: str | None = None
        self.current_version: str | None = None
        self.latest_version: str | None = None
        self.version_checking = False
        self.version_check_error: str | None = None
        self.paused = False
        self.updating = False

    def start(self, notify: bool = False) -> None:
        with self.lock:
            if self.is_running():
                return

            self.last_start_time = time.time()
            if not PLEXTRAKTSYNC_PYTHON.exists():
                self.last_error = f"Missing watcher Python at {PLEXTRAKTSYNC_PYTHON}"
                raise FileNotFoundError(self.last_error)

            # Keep PlexTraktSync in its own pipx-managed environment. The tray app
            # can be packaged independently without bundling PlexTraktSync itself.
            watcher_env = os.environ.copy()
            watcher_env.update(
                {
                    "PYTHONIOENCODING": "utf-8",
                    "PYTHONUTF8": "1",
                    "NO_COLOR": "1",
                    "TERM": "dumb",
                }
            )
            self.process = subprocess.Popen(
                [str(PLEXTRAKTSYNC_PYTHON), "-m", "plextraktsync", "watch"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                env=watcher_env,
                creationflags=CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP,
            )
            self.last_error = None
            self.last_connected_at = self.last_start_time
            self.auto_restart = True
            self.paused = False
            if notify:
                notify_message("Watcher started.")

    def stop(self, notify: bool = False) -> None:
        with self.lock:
            process = self.process

        if process is None:
            return

        if process.poll() is None:
            try:
                process.send_signal(signal.CTRL_BREAK_EVENT)
                process.wait(timeout=10)
            except (subprocess.TimeoutExpired, OSError):
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=5)

        with self.lock:
            if self.process is process:
                self.process = None

        if notify:
            notify_message("Watcher stopped.")

    def restart(self) -> None:
        self.stop()
        time.sleep(1)
        self.start(notify=True)

    def stop_manually(self) -> None:
        self.paused = False
        self.auto_restart = False
        self.stop(notify=True)

    def pause(self) -> None:
        self.paused = True
        self.auto_restart = False
        self.stop(notify=True)

    def resume(self) -> None:
        self.paused = False
        self.auto_restart = True
        self.start(notify=True)

    def is_running(self) -> bool:
        return self.process is not None and self.process.poll() is None

    def exit_code(self) -> int | None:
        if self.process is None:
            return None
        return self.process.poll()

    def status_text(self) -> str:
        if self.is_running():
            return "Watcher: running"
        if self.paused:
            return "Watcher: paused"
        if not self.auto_restart:
            return "Watcher: stopped"
        if self.last_error:
            return f"Watcher error: {self.last_error}"
        code = self.exit_code()
        if code is None:
            return "Watcher: stopped"
        return f"Watcher exited: {code}"

    def connected_text(self) -> str:
        if self.last_connected_at is None:
            return "Last connect: unknown"
        stamp = datetime.fromtimestamp(self.last_connected_at).strftime("%Y-%m-%d %I:%M:%S %p")
        return f"Last connect: {stamp}"

    def update_text(self) -> str:
        if self.updating:
            return "Update: running"
        if self.version_checking:
            return "PlexTraktSync: checking"
        if self.current_version and self.latest_version:
            if self.current_version == self.latest_version:
                return f"PlexTraktSync: current ({self.current_version})"
            return f"PlexTraktSync: update {self.current_version} -> {self.latest_version}"
        if self.version_check_error:
            return "PlexTraktSync: update status unknown"
        if self.last_update_result:
            return f"Update: {self.last_update_result}"
        return "Update: not checked"

    def update_action_text(self) -> str:
        if self.updating:
            return "Updating PlexTraktSync..."
        if self.version_checking:
            return "Checking for PlexTraktSync Update..."
        if self.current_version and self.latest_version and self.current_version == self.latest_version:
            return f"PlexTraktSync Current ({self.current_version})"
        if self.current_version and self.latest_version and self.current_version != self.latest_version:
            return "Install PlexTraktSync Update"
        return "Check for PlexTraktSync Update"

    def check_versions(self, notify: bool = False, already_claimed: bool = False) -> None:
        with self.lock:
            if self.version_checking and not already_claimed:
                return
            if not already_claimed:
                self.version_checking = True
            self.version_check_error = None

        try:
            current_version = get_installed_plextraktsync_version()
            latest_version = get_latest_plextraktsync_version()
            with self.lock:
                self.current_version = current_version
                self.latest_version = latest_version
            if notify:
                notify_message(self.update_text())
        except Exception as exc:
            with self.lock:
                self.version_check_error = str(exc)
            if notify:
                notify_message(f"Version check failed: {exc}")
        finally:
            with self.lock:
                self.version_checking = False
            refresh_icon()

    def upgrade_plextraktsync(self, already_claimed: bool = False) -> None:
        with self.lock:
            if self.updating and not already_claimed:
                return
            if not already_claimed:
                self.updating = True
        self.last_update_result = None
        was_paused = self.paused
        previous_auto_restart = self.auto_restart
        self.auto_restart = False

        try:
            self.stop()

            pipx = pipx_command()

            result = subprocess.run(
                [pipx, "upgrade", "plextraktsync"],
                capture_output=True,
                text=True,
                creationflags=CREATE_NO_WINDOW,
                timeout=300,
            )
            output = (result.stdout + result.stderr).strip()
            if result.returncode != 0:
                raise RuntimeError(output or f"pipx exited with {result.returncode}")

            self.last_update_result = "complete"
            self.check_versions()
            notify_message("PlexTraktSync update complete.")
        except Exception as exc:
            self.last_update_result = "failed"
            self.last_error = str(exc)
            notify_message(f"PlexTraktSync update failed: {exc}")
        finally:
            with self.lock:
                self.updating = False
                self.paused = was_paused
                self.auto_restart = previous_auto_restart
            if not self.paused and self.auto_restart:
                try:
                    self.start()
                except Exception as exc:
                    self.last_error = str(exc)
            refresh_icon()

    def claim_update_action(self):
        with self.lock:
            if self.updating or self.version_checking:
                return None, None
            if self.current_version and self.latest_version and self.current_version != self.latest_version:
                self.updating = True
                return "Updating PlexTraktSync...", lambda: self.upgrade_plextraktsync(already_claimed=True)
            self.version_checking = True
            return "Checking PlexTraktSync version...", lambda: self.check_versions(notify=True, already_claimed=True)


manager = WatcherManager()
completed_movie_sync = CompletedMovieSync(LOG_FILE, COMPLETED_MOVIE_STATE_FILE)
tray_icon: pystray.Icon | None = None
shutdown_event = threading.Event()
instance_mutex = None
_last_icon_state: tuple[bool, bool, str | None, bool | None] | None = None


class StartupCache:
    def __init__(self) -> None:
        self.value: bool | None = None
        self.checked_at = 0.0
        self.lock = threading.Lock()

    def get(self, ttl: float = 60.0) -> bool:
        now = time.time()
        with self.lock:
            if self.value is None or now - self.checked_at > ttl:
                self.value = startup_enabled()
                self.checked_at = now
            return self.value

    def invalidate(self) -> None:
        with self.lock:
            self.value = None


startup_cache = StartupCache()


def create_image(color: str) -> Image.Image:
    image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((6, 6, 58, 58), radius=14, fill=color)
    draw.rounded_rectangle((12, 12, 52, 52), radius=10, outline="white", width=2)
    draw.rectangle((22, 20, 30, 44), fill="white")
    draw.polygon([(30, 20), (45, 32), (30, 44)], fill="white")
    draw.ellipse((42, 10, 54, 22), fill="#facc15")
    return image


def current_icon_image() -> Image.Image:
    if auth_health.trakt_ok is False:
        return create_image("#9a6700")
    if manager.is_running():
        return create_image("#18794e")
    if manager.paused:
        return create_image("#9a6700")
    return create_image("#b42318")


def notify_message(message: str) -> None:
    if tray_icon is None:
        return
    try:
        tray_icon.notify(message, APP_NAME)
    except Exception:
        pass


def pipx_command() -> str:
    discovered = shutil.which("pipx")
    if discovered:
        return discovered

    local_bin_candidate = Path.home() / ".local" / "bin" / "pipx.exe"
    if local_bin_candidate.exists():
        return str(local_bin_candidate)

    python_scripts_dir = Path(os.environ["APPDATA"]) / "Python"
    for candidate in sorted(python_scripts_dir.glob("Python*\\Scripts\\pipx.exe"), reverse=True):
        if candidate.exists():
            return str(candidate)

    raise FileNotFoundError("pipx.exe was not found on PATH or in the expected Python Scripts folder")


def get_installed_plextraktsync_version() -> str:
    result = subprocess.run(
        [str(PLEXTRAKTSYNC_PYTHON), "-m", "pip", "show", "plextraktsync"],
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
        timeout=30,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Unable to read installed PlexTraktSync version")

    for line in result.stdout.splitlines():
        if line.lower().startswith("version:"):
            return line.split(":", 1)[1].strip()

    raise RuntimeError("Installed PlexTraktSync version was not found")


def get_latest_plextraktsync_version() -> str:
    with urllib.request.urlopen(PYPI_JSON_URL, timeout=15) as response:
        payload = json.loads(response.read().decode("utf-8"))
    version = payload.get("info", {}).get("version")
    if not version:
        raise RuntimeError("Latest PlexTraktSync version was not found on PyPI")
    return str(version)


def plextraktsync_server_name() -> str:
    env_path = LOCAL_APPDATA / ".env"
    try:
        for line in env_path.read_text(encoding="utf-8").splitlines():
            if line.startswith("PLEX_SERVER="):
                return line.split("=", 1)[1].strip() or "default"
    except OSError:
        pass
    return "default"


def plextraktsync_server_connection() -> tuple[str, str]:
    """Return the Plex base URL and token from PlexTraktSync's working server config."""

    server_name = plextraktsync_server_name()
    try:
        text = SERVERS_CONFIG_FILE.read_text(encoding="utf-8")
    except OSError as exc:
        raise RuntimeError("PlexTraktSync servers.yml not found") from exc

    pattern = re.compile(rf"^  {re.escape(server_name)}:\n(?P<body>(?:    .+\n|    - .+\n?)*)", re.MULTILINE)
    match = pattern.search(text)
    if match is None and server_name != "default":
        match = re.search(r"^  default:\n(?P<body>(?:    .+\n|    - .+\n?)*)", text, re.MULTILINE)
    if match is None:
        raise RuntimeError(f"Plex server '{server_name}' not found in servers.yml")

    body = match.group("body")
    token_match = re.search(r"^\s+token:\s*(?P<token>\S+)", body, re.MULTILINE)
    if token_match is None:
        raise RuntimeError(f"Plex token missing for server '{server_name}'")
    token = token_match.group("token").strip()

    urls: list[str] = []
    in_urls = False
    for line in body.splitlines():
        stripped = line.strip()
        if stripped == "urls:":
            in_urls = True
            continue
        if in_urls and stripped.startswith("- "):
            value = stripped[2:].strip()
            if value and value != "null":
                urls.append(value.rstrip("/"))
            continue
        if in_urls and stripped and not stripped.startswith("- "):
            break

    for url in urls:
        try:
            query = urllib.parse.urlencode({"X-Plex-Token": token})
            with urllib.request.urlopen(f"{url}/?{query}", timeout=8):
                return url, token
        except Exception:
            continue

    if urls:
        return urls[0], token
    raise RuntimeError(f"Plex URL missing for server '{server_name}'")


def plex_metadata_root(rating_key: str) -> ET.Element:
    candidates = [plextraktsync_server_connection()]
    last_error: Exception | None = None
    for base_url, token in candidates:
        try:
            query = urllib.parse.urlencode({"X-Plex-Token": token})
            url = f"{base_url}/library/metadata/{rating_key}?{query}"
            with urllib.request.urlopen(url, timeout=20) as response:
                return ET.fromstring(response.read())
        except Exception as exc:
            last_error = exc

    raise RuntimeError(f"Unable to read Plex item {rating_key}: {last_error}")


def completed_movie_event_from_log_line(line: str) -> WatchedMovieEvent | None:
    match = WATCHED_MOVIE_PATTERN.match(line)
    if not match or match.group("kind") != "Movie":
        return None

    try:
        progress = float(match.group("progress"))
    except ValueError:
        return None
    played = match.group("played") == "True"
    stopped_at_threshold = match.group("state").lower() == "stopped" and progress >= WATCHED_PROGRESS_THRESHOLD
    if not played and not stopped_at_threshold:
        return None

    try:
        timestamp = datetime.strptime(match.group("timestamp"), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None

    return WatchedMovieEvent(
        timestamp=timestamp,
        rating_key=match.group("rating_key"),
        title=match.group("title").strip(),
        progress=progress,
    )


def plex_movie_ids(rating_key: str) -> dict[str, object]:
    root = plex_metadata_root(rating_key)
    ids: dict[str, object] = {}
    for guid in root.findall(".//Guid"):
        value = guid.attrib.get("id", "")
        if value.startswith("imdb://"):
            ids["imdb"] = value.removeprefix("imdb://")
        elif value.startswith("tmdb://"):
            tmdb_id = value.removeprefix("tmdb://")
            if tmdb_id.isdigit():
                ids["tmdb"] = int(tmdb_id)
        elif value.startswith("tvdb://"):
            tvdb_id = value.removeprefix("tvdb://")
            if tvdb_id.isdigit():
                ids["tvdb"] = int(tvdb_id)
    if not ids:
        raise RuntimeError(f"No Trakt-compatible IDs found for Plex item {rating_key}")
    return ids


def trakt_headers() -> dict[str, str]:
    payload = json.loads(PYTRAKT_CONFIG_FILE.read_text(encoding="utf-8"))
    client_id = str(payload.get("CLIENT_ID", "")).strip()
    token = str(payload.get("OAUTH_TOKEN", "")).strip()
    if not client_id or not token:
        raise RuntimeError("Trakt token missing")
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "trakt-api-key": client_id,
        "trakt-api-version": "2",
        "User-Agent": APP_NAME,
    }


def mark_trakt_movie_watched(event: WatchedMovieEvent) -> None:
    ids = plex_movie_ids(event.rating_key)
    if trakt_movie_already_watched(ids, event.timestamp):
        return

    body = json.dumps(
        {
            "movies": [
                {
                    "watched_at": event.timestamp.astimezone().astimezone(timezone.utc).isoformat().replace("+00:00", "Z"),
                    "ids": ids,
                }
            ]
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        "https://api.trakt.tv/sync/history",
        data=body,
        headers=trakt_headers(),
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        if response.status not in {200, 201}:
            raise RuntimeError(f"HTTP {response.status}")


def trakt_movie_already_watched(ids: dict[str, object], watched_at: datetime) -> bool:
    request = urllib.request.Request(
        "https://api.trakt.tv/sync/history/movies?limit=50",
        headers=trakt_headers(),
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        history = json.loads(response.read().decode("utf-8"))

    for item in history:
        movie_ids = item.get("movie", {}).get("ids", {})
        if not movie_ids_match(ids, movie_ids):
            continue
        try:
            existing = datetime.fromisoformat(str(item.get("watched_at", "")).replace("Z", "+00:00"))
        except ValueError:
            continue
        if abs((existing.replace(tzinfo=None) - watched_at).total_seconds()) <= 36 * 60 * 60:
            return True
    return False


def movie_ids_match(left: dict[str, object], right: dict[str, object]) -> bool:
    for key in ("imdb", "tmdb", "tvdb"):
        if key in left and key in right and str(left[key]) == str(right[key]):
            return True
    return False


class AuthHealth:
    """Checks destination auth independently from the watcher process."""

    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.running = False
        self.last_trakt_check = 0.0
        self.trakt_status = "Trakt: not checked"
        self.trakt_ok: bool | None = None

    def trakt_text(self) -> str:
        with self.lock:
            return self.trakt_status

    def check_if_due(self, force: bool = False) -> None:
        now = time.time()
        retry_after = TRAKT_AUTH_CHECK_SECONDS if self.trakt_ok is not False else TRAKT_AUTH_RETRY_SECONDS
        trakt_due = force or now - self.last_trakt_check >= retry_after
        if not trakt_due:
            return

        with self.lock:
            if self.running:
                return
            self.running = True

        thread = threading.Thread(
            target=self._run_checks,
            args=(force,),
            daemon=True,
        )
        thread.start()

    def _run_checks(self, notify_success: bool) -> None:
        try:
            self._check_trakt(notify_success=notify_success)
        finally:
            with self.lock:
                self.running = False
            refresh_icon()

    def _check_trakt(self, notify_success: bool = False) -> None:
        self.last_trakt_check = time.time()
        try:
            payload = json.loads(PYTRAKT_CONFIG_FILE.read_text(encoding="utf-8"))
            client_id = str(payload.get("CLIENT_ID", "")).strip()
            token = str(payload.get("OAUTH_TOKEN", "")).strip()
            if not client_id or not token:
                raise RuntimeError("token missing")

            request = urllib.request.Request(
                TRAKT_API_SETTINGS_URL,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                    "trakt-api-key": client_id,
                    "trakt-api-version": "2",
                    "User-Agent": APP_NAME,
                },
            )
            with urllib.request.urlopen(request, timeout=20) as response:
                if response.status != 200:
                    raise RuntimeError(f"HTTP {response.status}")

            self._set_trakt(True, "Trakt: auth ok", notify_success=notify_success)
        except Exception as exc:
            self._set_trakt(False, f"Trakt auth failed: {friendly_error(exc)}")

    def _set_trakt(self, ok: bool, status: str, notify_success: bool = False) -> None:
        previous = self.trakt_ok
        with self.lock:
            self.trakt_ok = ok
            self.trakt_status = status
        if not ok and previous is not False:
            notify_message(f"{status}. Run PlexTraktSync trakt-login.")
        elif ok and notify_success:
            notify_message(status)


def friendly_error(exc: Exception) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        if exc.code in {401, 403}:
            return "unauthorized"
        return f"HTTP {exc.code}"
    if isinstance(exc, FileNotFoundError):
        return "token file missing"
    return type(exc).__name__


auth_health = AuthHealth()


def current_playback_text() -> str:
    """Return a short human status by reading recent PlexTraktSync watch logs."""

    if not LOG_FILE.exists():
        return "Running: idle"

    try:
        with LOG_FILE.open("rb") as log_file:
            log_file.seek(0, os.SEEK_END)
            size = log_file.tell()
            log_file.seek(max(0, size - LOG_TAIL_BYTES))
            lines = log_file.read().decode("utf-8", errors="replace").splitlines()
    except OSError:
        return "Running: unknown"

    # PlexTraktSync does not expose a status API, so the tray tooltip uses the
    # newest recent `on_play` log line and treats old/stopped entries as idle.
    for line in reversed(lines[-300:]):
        match = ON_PLAY_PATTERN.match(line)
        if not match:
            continue

        try:
            ts = datetime.strptime(match.group("timestamp"), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            ts = None

        if ts is not None and datetime.now() - ts > timedelta(minutes=PLAYBACK_STALE_MINUTES):
            return "Running: idle"

        state = match.group("state").lower()
        if state == "stopped":
            return "Running: idle"

        title = match.group("title").strip()
        try:
            progress = float(match.group("progress"))
        except ValueError:
            continue
        return f"Running: {title} [{state} {progress:.1f}%]"

    return "Running: idle"


def tooltip_text() -> str:
    return current_playback_text()


def run_powershell(command: str) -> subprocess.CompletedProcess[str]:
    """Run scheduled-task commands without opening a console window."""

    return subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            command,
        ],
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
    )


def startup_enabled() -> bool:
    result = run_powershell(
        f"(Get-ScheduledTask -TaskName '{TASK_NAME}' -ErrorAction Stop).Settings.Enabled"
    )
    return result.returncode == 0 and result.stdout.strip().lower() == "true"


def set_startup_enabled(enabled: bool) -> None:
    action = "Enable-ScheduledTask" if enabled else "Disable-ScheduledTask"
    result = run_powershell(f"{action} -TaskName '{TASK_NAME}'")
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip() or f"{action} failed"
        raise RuntimeError(message)


def refresh_icon() -> None:
    global _last_icon_state

    if tray_icon is None:
        return
    icon_state = (
        manager.is_running(),
        manager.paused,
        manager.last_error,
        auth_health.trakt_ok,
    )
    if icon_state != _last_icon_state:
        tray_icon.icon = current_icon_image()
        _last_icon_state = icon_state
    tray_icon.title = tooltip_text()
    tray_icon.update_menu()


def monitor_loop() -> None:
    """Refresh tray state and restart the watcher if it exits unexpectedly."""

    while not shutdown_event.is_set():
        try:
            if not manager.is_running() and manager.auto_restart and not manager.paused:
                if manager.last_start_time == 0 or (time.time() - manager.last_start_time) >= RESTART_DELAY_SECONDS:
                    try:
                        manager.start()
                    except Exception as exc:
                        manager.last_error = str(exc)
            auth_health.check_if_due()
            completed_movie_sync.check_log()
            refresh_icon()
        except Exception as exc:
            manager.last_error = str(exc)
            try:
                refresh_icon()
            except Exception:
                pass
        shutdown_event.wait(CHECK_INTERVAL_SECONDS)


def open_log() -> None:
    if LOG_FILE.exists():
        os.startfile(LOG_FILE)
    else:
        notify_message("Log file not found yet.")


def open_config_dir() -> None:
    os.startfile(str(LOCAL_APPDATA))


def open_plex_web() -> None:
    webbrowser.open(PLEX_WEB_URL)


def open_trakt_web() -> None:
    webbrowser.open(TRAKT_WEB_URL)


def on_start(_: pystray.Icon, __: Item) -> None:
    try:
        manager.paused = False
        manager.auto_restart = True
        manager.start(notify=True)
    except Exception as exc:
        manager.last_error = str(exc)
        notify_message(f"Start failed: {exc}")
    refresh_icon()


def on_stop(_: pystray.Icon, __: Item) -> None:
    manager.stop_manually()
    refresh_icon()


def on_restart(_: pystray.Icon, __: Item) -> None:
    try:
        manager.paused = False
        manager.auto_restart = True
        manager.restart()
    except Exception as exc:
        manager.last_error = str(exc)
        notify_message(f"Restart failed: {exc}")
    refresh_icon()


def on_open_log(_: pystray.Icon, __: Item) -> None:
    open_log()


def on_open_config(_: pystray.Icon, __: Item) -> None:
    open_config_dir()


def on_open_plex(_: pystray.Icon, __: Item) -> None:
    open_plex_web()


def on_open_trakt(_: pystray.Icon, __: Item) -> None:
    open_trakt_web()


def on_pause_resume(_: pystray.Icon, __: Item) -> None:
    if manager.paused:
        try:
            manager.resume()
        except Exception as exc:
            manager.last_error = str(exc)
            notify_message(f"Resume failed: {exc}")
    else:
        manager.pause()
    refresh_icon()


def on_toggle_startup(_: pystray.Icon, __: Item) -> None:
    try:
        enabled = startup_cache.get(ttl=0)
        set_startup_enabled(not enabled)
        startup_cache.invalidate()
        notify_message("Start with Windows enabled." if not enabled else "Start with Windows disabled.")
    except Exception as exc:
        notify_message(f"Startup toggle failed: {exc}")
    refresh_icon()


def on_update_plextraktsync(_: pystray.Icon, __: Item) -> None:
    message, target = manager.claim_update_action()
    if target is None:
        return
    notify_message(message)

    thread = threading.Thread(target=target, daemon=True)
    thread.start()
    refresh_icon()


def on_check_auth(_: pystray.Icon, __: Item) -> None:
    notify_message("Checking Trakt auth...")
    auth_health.check_if_due(force=True)
    refresh_icon()


def on_exit(icon: pystray.Icon, _: Item) -> None:
    shutdown_event.set()
    manager.stop()
    icon.stop()


def build_menu() -> pystray.Menu:
    return pystray.Menu(
        Item(lambda _: manager.status_text(), None, enabled=False),
        Item(lambda _: current_playback_text(), None, enabled=False),
        Item(lambda _: manager.connected_text(), None, enabled=False),
        Item(lambda _: manager.update_text(), None, enabled=False),
        Item(lambda _: auth_health.trakt_text(), None, enabled=False),
        Item(lambda _: completed_movie_sync.status_text(), None, enabled=False),
        Item(lambda _: f"Start with Windows: {'enabled' if startup_cache.get() else 'disabled'}", None, enabled=False),
        Item("Start Watcher", on_start, enabled=lambda _: not manager.is_running()),
        Item("Stop Watcher", on_stop, enabled=lambda _: manager.is_running()),
        Item(lambda _: "Resume Watcher" if manager.paused else "Pause Watcher", on_pause_resume),
        Item("Restart Watcher", on_restart),
        Item(lambda _: manager.update_action_text(), on_update_plextraktsync, enabled=lambda _: not manager.updating and not manager.version_checking),
        Item("Check Auth Now", on_check_auth, enabled=lambda _: not auth_health.running),
        Item(lambda _: "Disable Start With Windows" if startup_cache.get() else "Enable Start With Windows", on_toggle_startup),
        Item("Open Plex Web", on_open_plex),
        Item("Open Trakt", on_open_trakt),
        Item("Open Log", on_open_log),
        Item("Open Config Folder", on_open_config),
        Item("Exit", on_exit),
    )


def main() -> int:
    global tray_icon
    global instance_mutex

    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.CloseHandle.restype = ctypes.c_int
    kernel32.GetLastError.restype = ctypes.c_uint32
    # A named Windows mutex prevents duplicate tray apps when the scheduled task
    # is started manually while an instance is already running.
    instance_mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not instance_mutex:
        return 1
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(instance_mutex)
        return 0

    LOCAL_APPDATA.mkdir(parents=True, exist_ok=True)

    try:
        manager.start()
    except Exception as exc:
        manager.last_error = str(exc)

    threading.Thread(target=manager.check_versions, daemon=True).start()
    auth_health.check_if_due(force=True)

    tray_icon = pystray.Icon(APP_NAME, current_icon_image(), tooltip_text(), build_menu())

    monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
    monitor_thread.start()

    tray_icon.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
