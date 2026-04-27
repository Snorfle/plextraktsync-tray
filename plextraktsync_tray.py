import ctypes
import json
import os
import re
import shutil
import subprocess
import threading
import time
import urllib.request
import webbrowser
from datetime import datetime, timedelta
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
CHECK_INTERVAL_SECONDS = 10
RESTART_DELAY_SECONDS = 15
LOG_TAIL_BYTES = 65536
CREATE_NO_WINDOW = 0x08000000
ERROR_ALREADY_EXISTS = 183
MUTEX_NAME = "Local\\PlexTraktSyncTrayApp"
PLEX_WEB_URL = "http://127.0.0.1:32400/web"
TRAKT_WEB_URL = "https://trakt.tv/"
PYPI_JSON_URL = "https://pypi.org/pypi/plextraktsync/json"
PLAYBACK_STALE_MINUTES = 30
ON_PLAY_PATTERN = re.compile(
    r"^(?P<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+\s+INFO\[.*?\]:on_play: <[^>]+:(?P<title>.+)>: (?P<progress>[\d.]+)%, State: (?P<state>\w+)",
)


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
                creationflags=CREATE_NO_WINDOW,
            )
            self.last_error = None
            self.last_start_time = time.time()
            self.last_connected_at = self.last_start_time
            self.auto_restart = True
            self.paused = False
            if notify:
                notify_message("Watcher started.")

    def stop(self, notify: bool = False) -> None:
        with self.lock:
            process = self.process
            self.process = None

        if process is None:
            return

        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)

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

    def check_versions(self, notify: bool = False) -> None:
        if self.version_checking:
            return

        self.version_checking = True
        self.version_check_error = None

        try:
            self.current_version = get_installed_plextraktsync_version()
            self.latest_version = get_latest_plextraktsync_version()
            if notify:
                notify_message(self.update_text())
        except Exception as exc:
            self.version_check_error = str(exc)
            if notify:
                notify_message(f"Version check failed: {exc}")
        finally:
            self.version_checking = False
            refresh_icon()

    def upgrade_plextraktsync(self) -> None:
        if self.updating:
            return

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
            self.updating = False
            self.paused = was_paused
            self.auto_restart = previous_auto_restart
            if not self.paused and self.auto_restart:
                try:
                    self.start()
                except Exception as exc:
                    self.last_error = str(exc)
            refresh_icon()


manager = WatcherManager()
tray_icon: pystray.Icon | None = None
shutdown_event = threading.Event()
instance_mutex = None


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

        title = match.group("title").replace("-", " ").strip()
        progress = float(match.group("progress"))
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
    if tray_icon is None:
        return
    tray_icon.icon = current_icon_image()
    tray_icon.title = tooltip_text()
    tray_icon.update_menu()


def monitor_loop() -> None:
    """Refresh tray state and restart the watcher if it exits unexpectedly."""

    while not shutdown_event.is_set():
        if not manager.is_running() and manager.auto_restart and not manager.paused:
            if manager.last_start_time == 0 or (time.time() - manager.last_start_time) >= RESTART_DELAY_SECONDS:
                try:
                    manager.start()
                except Exception as exc:
                    manager.last_error = str(exc)
        refresh_icon()
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
        enabled = startup_enabled()
        set_startup_enabled(not enabled)
        notify_message("Start with Windows enabled." if not enabled else "Start with Windows disabled.")
    except Exception as exc:
        notify_message(f"Startup toggle failed: {exc}")
    refresh_icon()


def on_update_plextraktsync(_: pystray.Icon, __: Item) -> None:
    if manager.current_version and manager.latest_version and manager.current_version != manager.latest_version:
        notify_message("Updating PlexTraktSync...")
        target = manager.upgrade_plextraktsync
    else:
        notify_message("Checking PlexTraktSync version...")
        target = lambda: manager.check_versions(notify=True)

    thread = threading.Thread(target=target, daemon=True)
    thread.start()
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
        Item(lambda _: f"Start with Windows: {'enabled' if startup_enabled() else 'disabled'}", None, enabled=False),
        Item("Start Watcher", on_start, enabled=lambda _: not manager.is_running()),
        Item("Stop Watcher", on_stop, enabled=lambda _: manager.is_running()),
        Item(lambda _: "Resume Watcher" if manager.paused else "Pause Watcher", on_pause_resume),
        Item("Restart Watcher", on_restart),
        Item(lambda _: manager.update_action_text(), on_update_plextraktsync, enabled=lambda _: not manager.updating and not manager.version_checking),
        Item(lambda _: "Disable Start With Windows" if startup_enabled() else "Enable Start With Windows", on_toggle_startup),
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

    tray_icon = pystray.Icon(APP_NAME, current_icon_image(), tooltip_text(), build_menu())

    monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
    monitor_thread.start()

    tray_icon.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
