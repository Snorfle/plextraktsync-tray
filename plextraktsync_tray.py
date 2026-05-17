import ctypes
import json
import os
import re
import signal
import shutil
import sqlite3
import subprocess
import threading
import time
import urllib.parse
import urllib.error
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
from contextlib import contextmanager
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
LEGACY_COMPLETED_MOVIE_STATE_FILE = LOCAL_APPDATA / "completed_movie_sync_state.json"
TARGET_SYNC_LEDGER_FILE = LOCAL_APPDATA / "target_sync.sqlite"
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
WATCHED_MEDIA_PATTERN = re.compile(
    r"^(?P<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+\s+INFO\[.*?\]:on_play: "
    r"<(?P<kind>[^:>]+):(?P<rating_key>\d+):(?P<title>.+)>: (?P<progress>\d+(?:\.\d+)?)%, "
    r"State: (?P<state>\w+), Played: (?P<played>True|False)",
)
TERMINAL_TARGET_STATUSES = {
    "synced",
    "already_present",
    "planned_read_only",
    "blocked",
    "not_applicable",
}

@dataclass(frozen=True)
class PlexPlaybackEvent:
    timestamp: datetime
    kind: str
    rating_key: str
    title: str
    progress: float


@dataclass(frozen=True)
class MediaEvent:
    source_event_key: str
    content_type: str
    title: str
    watched_at: datetime
    source_item_key: str
    progress: float
    tmdb_id: int | None = None
    imdb_id: str | None = None
    tvdb_id: int | None = None
    season_number: int | None = None
    episode_number: int | None = None
    episode_title: str | None = None


LEDGER_SCHEMA = """
create table if not exists media_events(
  id integer primary key,
  source text not null,
  source_event_key text not null unique,
  content_type text not null,
  title text not null,
  season_number integer,
  episode_number integer,
  episode_title text,
  watched_at text not null,
  tmdb_id integer,
  imdb_id text,
  tvdb_id integer,
  source_item_key text not null,
  progress real not null,
  created_at text not null
);

create table if not exists target_attempts(
  id integer primary key,
  media_event_id integer not null references media_events(id) on delete cascade,
  target text not null,
  status text not null,
  request_summary text,
  response_summary text,
  attempted_at text not null
);
"""

class TargetLedger:
    """Durable per-target sync state for Plex playback events."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.lock = threading.Lock()
        self.init()

    def connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        conn.execute("pragma foreign_keys = on")
        return conn

    @contextmanager
    def connection(self):
        conn = self.connect()
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def init(self) -> None:
        with self.lock:
            with self.connection() as conn:
                conn.executescript(LEDGER_SCHEMA)

    def upsert_media_event(self, event: MediaEvent) -> int:
        with self.lock:
            with self.connection() as conn:
                existing = conn.execute(
                    "select id from media_events where source_event_key = ?",
                    (event.source_event_key,),
                ).fetchone()
                if existing:
                    return int(existing["id"])

                cursor = conn.execute(
                    """
                    insert into media_events(
                      source, source_event_key, content_type, title, season_number,
                      episode_number, episode_title, watched_at, tmdb_id, imdb_id,
                      tvdb_id, source_item_key, progress, created_at
                    ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "plex",
                        event.source_event_key,
                        event.content_type,
                        event.title,
                        event.season_number,
                        event.episode_number,
                        event.episode_title,
                        event.watched_at.isoformat(timespec="seconds"),
                        event.tmdb_id,
                        event.imdb_id,
                        event.tvdb_id,
                        event.source_item_key,
                        event.progress,
                        datetime.now().isoformat(timespec="seconds"),
                    ),
                )
                return int(cursor.lastrowid)

    def media_event_id(self, source_event_key: str) -> int | None:
        with self.lock:
            with self.connection() as conn:
                row = conn.execute(
                    "select id from media_events where source_event_key = ?",
                    (source_event_key,),
                ).fetchone()
        return int(row["id"]) if row else None

    def target_confirmed(self, media_event_id: int, target: str) -> bool:
        statuses = tuple(TERMINAL_TARGET_STATUSES)
        placeholders = ", ".join("?" for _ in statuses)
        with self.lock:
            with self.connection() as conn:
                row = conn.execute(
                    f"""
                    select 1
                      from target_attempts
                     where media_event_id = ?
                       and target = ?
                       and status in ({placeholders})
                     limit 1
                    """,
                    (media_event_id, target, *statuses),
                ).fetchone()
        return row is not None

    def record_target_attempt(
        self,
        media_event_id: int,
        target: str,
        status: str,
        request_summary: str | None = None,
        response_summary: str | None = None,
    ) -> None:
        with self.lock:
            with self.connection() as conn:
                conn.execute(
                    """
                    insert into target_attempts(
                      media_event_id, target, status, request_summary,
                      response_summary, attempted_at
                    ) values (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        media_event_id,
                        target,
                        status,
                        request_summary,
                        response_summary,
                        datetime.now().isoformat(timespec="seconds"),
                    ),
                )

    def target_status_text(self) -> str:
        with self.lock:
            with self.connection() as conn:
                row = conn.execute(
                    """
                    select
                      count(*) as total,
                      sum(case when content_type = 'movie' then 1 else 0 end) as movies,
                      sum(case when content_type = 'episode' then 1 else 0 end) as episodes
                    from media_events
                    """
                ).fetchone()
        total = int(row["total"] or 0)
        movies = int(row["movies"] or 0)
        episodes = int(row["episodes"] or 0)
        return f"Ledger: {total} events ({movies} movies, {episodes} episodes)"

    def target_summary_text(self) -> str:
        with self.lock:
            with self.connection() as conn:
                rows = list(
                    conn.execute(
                        """
                        select target, status, count(*) as count
                          from target_attempts
                         group by target, status
                         order by target, status
                        """
                    )
                )
        if not rows:
            return "Targets: no attempts yet"
        parts = [f"{row['target']} {row['status']}={row['count']}" for row in rows]
        return "Targets: " + ", ".join(parts[:4])


class SyncTarget:
    name = "target"

    def is_configured(self) -> bool:
        return True

    def applies_to(self, event: MediaEvent) -> bool:
        return True

    def status_text(self) -> str:
        return f"{self.name}: ready"

    def sync(self, event: MediaEvent) -> str:
        raise NotImplementedError


class TraktTarget(SyncTarget):
    name = "trakt"

    def applies_to(self, event: MediaEvent) -> bool:
        # Keep current public behavior: the tray only fills missed movie history.
        return event.content_type == "movie"

    def sync(self, event: MediaEvent) -> str:
        return mark_trakt_movie_watched(event)

    def status_text(self) -> str:
        return auth_health.trakt_text()


class SerializdTarget(SyncTarget):
    name = "serializd"

    def is_configured(self) -> bool:
        return True

    def applies_to(self, event: MediaEvent) -> bool:
        return event.content_type == "episode"

    def status_text(self) -> str:
        return "Serializd: read-only planning"

    def sync(self, event: MediaEvent) -> str:
        if event.tmdb_id is None:
            return "blocked"
        if event.season_number is None or event.episode_number is None:
            return "blocked"
        return "planned_read_only"


class TargetDispatcher:
    def __init__(self, targets: list[SyncTarget]) -> None:
        self.targets = targets

    def configured_targets_for(self, event: MediaEvent) -> list[SyncTarget]:
        return [target for target in self.targets if target.applies_to(event) and target.is_configured()]

    def may_have_pending_work(self, event: PlexPlaybackEvent, media_event_id: int | None, ledger: TargetLedger) -> bool:
        if media_event_id is None:
            return True
        content_type = "movie" if event.kind.lower() == "movie" else "episode"
        probe = MediaEvent(
            source_event_key=plex_source_event_key(event),
            content_type=content_type,
            title=event.title,
            watched_at=event.timestamp,
            source_item_key=f"plex:{event.rating_key}",
            progress=event.progress,
        )
        return any(
            not ledger.target_confirmed(media_event_id, target.name)
            for target in self.configured_targets_for(probe)
        )

    def status_lines(self) -> list[str]:
        return [target.status_text() for target in self.targets if target.name != "trakt"]


class CompletedMediaSync:
    """Records completed Plex playback events and runs the Trakt movie fallback."""

    def __init__(self, log_path: Path, ledger: TargetLedger, dispatcher: TargetDispatcher, legacy_state_path: Path) -> None:
        self.log_path = log_path
        self.ledger = ledger
        self.dispatcher = dispatcher
        self.legacy_state_path = legacy_state_path
        self.lock = threading.Lock()
        self.running = False
        self.last_status = "Target sync: waiting"
        self.legacy_synced_keys = self._load_legacy_synced_keys()

    def status_text(self) -> str:
        with self.lock:
            running = self.running
            last_status = self.last_status
        if running:
            return "Target sync: processing"
        return last_status

    def check_log(self) -> None:
        with self.lock:
            if self.running:
                return
        event = self._latest_unsynced_event()
        if event is None:
            self.last_status = "Target sync: waiting"
            return

        thread = threading.Thread(target=self._sync_event, args=(event,), daemon=True)
        thread.start()

    def _latest_unsynced_event(self) -> PlexPlaybackEvent | None:
        if not self.log_path.exists():
            return None

        try:
            with self.log_path.open("rb") as log_file:
                log_file.seek(0, os.SEEK_END)
                size = log_file.tell()
                log_file.seek(max(0, size - LOG_TAIL_BYTES))
                lines = log_file.read().decode("utf-8", errors="replace").splitlines()
        except OSError:
            self.last_status = "Target sync: log unavailable"
            return None

        for line in reversed(lines[-500:]):
            event = completed_media_event_from_log_line(line)
            if event is None:
                continue
            source_event_key = plex_source_event_key(event)
            media_event_id = self.ledger.media_event_id(source_event_key)
            if event.kind.lower() == "movie":
                if self._legacy_sync_key(event) in self.legacy_synced_keys:
                    continue
                if self.dispatcher.may_have_pending_work(event, media_event_id, self.ledger):
                    return event
            elif self.dispatcher.may_have_pending_work(event, media_event_id, self.ledger):
                return event
        return None

    def _sync_event(self, event: PlexPlaybackEvent) -> None:
        with self.lock:
            if self.running:
                return
            self.running = True

        media_event_id: int | None = None
        current_target = "target"
        try:
            self.last_status = f"Target sync: recording {event.title}"
            media_event = media_event_from_plex_event(event)
            media_event_id = self.ledger.upsert_media_event(media_event)

            targets = self.dispatcher.configured_targets_for(media_event)
            if not targets:
                self.last_status = f"Target sync: recorded {media_event.title}"
                return

            for target in targets:
                if self.ledger.target_confirmed(media_event_id, target.name):
                    continue
                current_target = target.name
                self.last_status = f"{target.name.title()}: syncing {media_event.title}"
                status = target.sync(media_event)
                self.ledger.record_target_attempt(
                    media_event_id,
                    target.name,
                    status,
                    request_summary=f"{media_event.content_type} {media_event.title}",
                    response_summary=f"{target.name} returned {status}",
                )
                self.last_status = f"{target.name.title()}: {status.replace('_', ' ')} {media_event.title}"
                if target.name == "trakt" and status == "synced":
                    notify_message(f"Marked {media_event.title} watched on Trakt.")
        except Exception as exc:
            if media_event_id is not None:
                self.ledger.record_target_attempt(
                    media_event_id,
                    current_target,
                    "failed_retryable",
                    request_summary=event.title,
                    response_summary=friendly_error(exc),
                )
            self.last_status = f"Target sync failed: {friendly_error(exc)}"
            notify_message(f"Trakt watched fallback failed: {friendly_error(exc)}.")
        finally:
            with self.lock:
                self.running = False
            refresh_icon()

    def _load_legacy_synced_keys(self) -> set[str]:
        if not self.legacy_state_path.exists():
            return set()
        try:
            payload = json.loads(self.legacy_state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return set()
        return {str(item) for item in payload.get("synced", [])}

    @staticmethod
    def _legacy_sync_key(event: PlexPlaybackEvent) -> str:
        return f"{event.rating_key}:{event.timestamp.isoformat(timespec='minutes')}"


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
target_ledger = TargetLedger(TARGET_SYNC_LEDGER_FILE)
trakt_target = TraktTarget()
serializd_target = SerializdTarget()
target_dispatcher = TargetDispatcher([trakt_target, serializd_target])
completed_media_sync = CompletedMediaSync(LOG_FILE, target_ledger, target_dispatcher, LEGACY_COMPLETED_MOVIE_STATE_FILE)
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


def completed_media_event_from_log_line(line: str) -> PlexPlaybackEvent | None:
    match = WATCHED_MEDIA_PATTERN.match(line)
    if not match or match.group("kind") not in {"Movie", "Episode"}:
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

    return PlexPlaybackEvent(
        timestamp=timestamp,
        kind=match.group("kind"),
        rating_key=match.group("rating_key"),
        title=match.group("title").strip(),
        progress=progress,
    )


def plex_source_event_key(event: PlexPlaybackEvent) -> str:
    return f"plex:{event.kind.lower()}:{event.rating_key}:{event.timestamp.isoformat(timespec='minutes')}"


def plex_video_node(root: ET.Element) -> ET.Element:
    if root.tag == "Video":
        return root
    video = root.find(".//Video")
    if video is None:
        raise RuntimeError("Plex metadata response did not include a Video item")
    return video


def plex_ids_from_root(root: ET.Element) -> dict[str, object]:
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
    return ids


def media_event_from_plex_event(event: PlexPlaybackEvent) -> MediaEvent:
    root = plex_metadata_root(event.rating_key)
    video = plex_video_node(root)
    kind = event.kind.lower()
    ids = plex_ids_from_root(root)

    title = video.attrib.get("title") or event.title
    season_number: int | None = None
    episode_number: int | None = None
    episode_title: str | None = None

    if kind == "movie":
        content_type = "movie"
    elif kind == "episode":
        content_type = "episode"
        episode_title = title
        title = video.attrib.get("grandparentTitle") or event.title
        season_number = int_or_none(video.attrib.get("parentIndex"))
        episode_number = int_or_none(video.attrib.get("index"))
        show_rating_key = video.attrib.get("grandparentRatingKey")
        if show_rating_key:
            try:
                ids = plex_ids_from_root(plex_metadata_root(show_rating_key)) or ids
            except Exception:
                pass
    else:
        raise RuntimeError(f"Unsupported Plex media kind {event.kind}")

    return MediaEvent(
        source_event_key=plex_source_event_key(event),
        content_type=content_type,
        title=title,
        watched_at=event.timestamp,
        source_item_key=f"plex:{event.rating_key}",
        progress=event.progress,
        tmdb_id=int_or_none(ids.get("tmdb")),
        imdb_id=str(ids["imdb"]) if ids.get("imdb") else None,
        tvdb_id=int_or_none(ids.get("tvdb")),
        season_number=season_number,
        episode_number=episode_number,
        episode_title=episode_title,
    )


def int_or_none(value: object) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


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


def mark_trakt_movie_watched(event: MediaEvent) -> str:
    ids = media_event_ids(event)
    if trakt_movie_already_watched(ids, event.watched_at):
        return "already_present"

    body = json.dumps(
        {
            "movies": [
                {
                    "watched_at": event.watched_at.astimezone().astimezone(timezone.utc).isoformat().replace("+00:00", "Z"),
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
    return "synced"


def media_event_ids(event: MediaEvent) -> dict[str, object]:
    ids: dict[str, object] = {}
    if event.imdb_id:
        ids["imdb"] = event.imdb_id
    if event.tmdb_id is not None:
        ids["tmdb"] = event.tmdb_id
    if event.tvdb_id is not None:
        ids["tvdb"] = event.tvdb_id
    if not ids:
        raise RuntimeError(f"No Trakt-compatible IDs found for {event.title}")
    return ids


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
            completed_media_sync.check_log()
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
        Item(lambda _: serializd_target.status_text(), None, enabled=False),
        Item(lambda _: completed_media_sync.status_text(), None, enabled=False),
        Item(lambda _: target_ledger.target_status_text(), None, enabled=False),
        Item(lambda _: target_ledger.target_summary_text(), None, enabled=False),
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
