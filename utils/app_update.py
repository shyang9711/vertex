import json
import os
import ssl
import sys
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk

# Optional logger if your app has one; safe fallback
try:
    from functions.utils.app_logging import get_logger
    LOG = get_logger("app_update")
except Exception:
    from utils.app_logging import get_logger
    LOG = None


def _parse_tag(tag_name: str | None) -> tuple[str | None, bool]:
    """
    Returns (version_without_prefix_or_suffix, is_forced)

    Supported GitHub release tag formats:
      - v0.1.55    -> ("0.1.55", False)  [minor update]
      - v0.1.55f   -> ("0.1.55", True)   [major/forced update]
      - 0.1.55f    -> ("0.1.55", True)
    """
    if not tag_name:
        return (None, False)

    tag = str(tag_name).strip()
    if tag.lower().startswith("v"):
        tag = tag[1:].strip()

    forced = False
    if tag.lower().endswith("f"):
        forced = True
        tag = tag[:-1].strip()

    return (tag or None, forced)


def _parse_version(v: str) -> tuple[int, int, int]:
    """
    "0.1.55" -> (0, 1, 55)
    Missing parts become 0.
    """
    parts = (v or "").strip().split(".")
    nums = []
    for p in parts[:3]:
        try:
            nums.append(int(p))
        except Exception:
            nums.append(0)
    while len(nums) < 3:
        nums.append(0)
    return (nums[0], nums[1], nums[2])


def _is_newer_version(candidate: str, current: str) -> bool:
    """
    True if candidate > current (semver-ish numeric compare).
    """
    return _parse_version(candidate) > _parse_version(current)


def _fetch_latest_release_json_silent(github_api_latest: str, app_name: str) -> dict | None:
    """
    Fetch latest release JSON from GitHub, silent (no UI).
    """
    try:
        req = urllib.request.Request(
            github_api_latest,
            headers={"User-Agent": app_name or "Vertex"},
        )
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, context=ctx, timeout=6) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def _latest_tag_from_release_json(release_json: dict) -> str | None:
    if not isinstance(release_json, dict):
        return None
    return str(release_json.get("tag_name") or "").strip() or None


def is_major_update_required_by_tag(
    *,
    app_version: str,
    github_api_latest: str,
    app_name: str,
) -> tuple[bool, str | None, str | None, bool]:
    """
    Returns (required, reason, latest_version, forced_flag)

    required becomes True ONLY when:
      - latest tag ends with 'f' (forced), AND
      - latest_version (without 'f') is newer than app_version
    """
    release_json = _fetch_latest_release_json_silent(github_api_latest, app_name)
    tag_name = _latest_tag_from_release_json(release_json)
    latest_version, forced = _parse_tag(tag_name)

    if LOG:
        LOG.info("latest_tag=%s latest_version=%s forced=%s app_version=%s", tag_name, latest_version, forced, app_version)

    if not latest_version:
        return (False, None, None, False)

    if forced and _is_newer_version(latest_version, app_version):
        reason = (
            "A required update is available.\n\n"
            f"Current version: {app_version}\n"
            f"Required version: {latest_version}\n"
        )
        return (True, reason, latest_version, True)

    return (False, None, latest_version, forced)


def enforce_major_update_on_startup(
    root: tk.Tk,
    *,
    app_name: str,
    app_version: str,
    github_api_latest: str,
    github_releases_url: str,
    update_asset_name: str = "vertex.exe",
) -> bool:
    """
    If the latest tag is forced (endswith 'f') and newer than current,
    show a blocking modal and prevent app usage.
    Returns True if blocked, else False.
    """
    required, reason, latest_version, forced = is_major_update_required_by_tag(
        app_version=app_version,
        github_api_latest=github_api_latest,
        app_name=app_name,
    )

    if not required or not reason:
        return False

    dlg = tk.Toplevel(root)
    dlg.title("Update required")
    dlg.resizable(False, False)
    dlg.transient(root)
    dlg.grab_set()

    def _ignore_close():
        pass

    dlg.protocol("WM_DELETE_WINDOW", _ignore_close)

    frm = ttk.Frame(dlg, padding=18)
    frm.pack(fill="both", expand=True)

    msg = (
        "Update required\n\n"
        f"Current version: {app_version}\n"
        f"Latest version:  {latest_version or 'unknown'}\n\n"
        f"{reason}\n"
        "Please update to continue."
    )
    ttk.Label(frm, text=msg, justify="left").pack(anchor="w")

    btns = ttk.Frame(frm)
    btns.pack(fill="x", pady=(16, 0))

    def _do_update():
        try:
            check_for_updates(
                parent=root,
                app_name=app_name,
                app_version=app_version,
                github_api_latest=github_api_latest,
                github_releases_url=github_releases_url,
                update_asset_name=update_asset_name,
            )
        finally:
            try:
                root.destroy()
            except Exception:
                pass

    def _exit():
        try:
            root.destroy()
        except Exception:
            pass

    ttk.Button(btns, text="Update", command=_do_update).pack(side="right", padx=(8, 0))
    ttk.Button(btns, text="Exit", command=_exit).pack(side="right")

    dlg.update_idletasks()
    w, h = dlg.winfo_width(), dlg.winfo_height()
    x = root.winfo_rootx() + (root.winfo_width() // 2) - (w // 2)
    y = root.winfo_rooty() + (root.winfo_height() // 2) - (h // 2)
    dlg.geometry(f"+{max(0, x)}+{max(0, y)}")

    return True


def check_for_updates(
    parent: tk.Misc | None,
    *,
    app_name: str,
    app_version: str,
    github_api_latest: str,
    github_releases_url: str,
    update_asset_name: str = "vertex.exe",
):
    """
    Manual update (Help -> Check for updates).
    - If running from source: open Releases page
    - If running frozen EXE: download EXE asset to <exe>.new and run update cmd to swap/relaunch
    """
    try:
        req = urllib.request.Request(github_api_latest, headers={"User-Agent": app_name})
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, context=ctx, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            messagebox.showinfo(
                "Updates",
                "No releases found on GitHub yet.\n\n"
                "Once you create a release for this app, the update checker will compare versions.",
                parent=parent,
            )
            return
        messagebox.showerror("Update Check Failed", f"HTTP error from GitHub:\n{e}", parent=parent)
        return
    except Exception as e:
        messagebox.showerror("Update Check Failed", f"Could not contact GitHub:\n{e}", parent=parent)
        return

    raw_tag = str(data.get("tag_name") or "").strip()
    latest_version, _forced = _parse_tag(raw_tag)

    if not latest_version:
        messagebox.showinfo(
            "Updates",
            "Latest release has no tag_name.\nUse tags like v0.1.55 or v0.1.55f on GitHub releases.",
            parent=parent,
        )
        return

    if not _is_newer_version(latest_version, app_version):
        messagebox.showinfo(
            "Up to date",
            f"You are running version {app_version}, which is the latest release.",
            parent=parent,
        )
        return

    if not getattr(sys, "frozen", False):
        if messagebox.askyesno(
            "Update available",
            f"Current version: {app_version}\nLatest version: {latest_version}\n\nOpen the GitHub Releases page?",
            parent=parent,
        ):
            webbrowser.open(github_releases_url)
        return

    assets = data.get("assets") or []
    exe_name = os.path.basename(sys.executable)
    url = None
    expected_size = None

    for a in assets:
        if a.get("name") == exe_name:
            url = a.get("browser_download_url")
            expected_size = a.get("size")
            break

    if not url and update_asset_name:
        for a in assets:
            if a.get("name") == update_asset_name:
                url = a.get("browser_download_url")
                expected_size = a.get("size")
                break

    if not url:
        if messagebox.askyesno(
            "Update available",
            f"Current version: {app_version}\nLatest version: {latest_version}\n\n"
            "No EXE asset found in the release.\nOpen Releases page in your browser?",
            parent=parent,
        ):
            webbrowser.open(github_releases_url)
        return

    app_folder = Path(sys.executable).resolve().parent
    dest = app_folder / f"{exe_name}.new"

    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={"User-Agent": "Vertex-Updater"})
        with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
            chunk_size = 1024 * 64
            with open(dest, "wb") as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                f.flush()
                os.fsync(f.fileno())

        actual_size = dest.stat().st_size
        if expected_size is not None and actual_size != int(expected_size):
            try:
                dest.unlink(missing_ok=True)
            except Exception:
                pass
            messagebox.showerror(
                "Update Download Failed",
                f"Downloaded file size mismatch.\n\nExpected: {expected_size} bytes\nGot: {actual_size} bytes\n\nPlease try Update again.",
                parent=parent,
            )
            return

        with open(dest, "rb") as f:
            sig = f.read(2)
        if sig != b"MZ":
            try:
                dest.unlink(missing_ok=True)
            except Exception:
                pass
            messagebox.showerror(
                "Update Download Failed",
                "Downloaded file is not a valid Windows executable.\n\nPlease try Update again.",
                parent=parent,
            )
            return

    except Exception as e:
        try:
            dest.unlink(missing_ok=True)
        except Exception:
            pass
        messagebox.showerror("Update Download Failed", f"Could not download the latest EXE.\n\n{e}", parent=parent)
        return

    import textwrap

    updater = app_folder / "update_vertex.cmd"

    cmd = textwrap.dedent(fr"""
        @echo off
        setlocal EnableExtensions
        echo Updating {app_name}.

        set "DIR=%~dp0"
        pushd "%DIR%" >nul 2>&1 || goto :fail

        if not defined LOCALAPPDATA set "LOCALAPPDATA=%USERPROFILE%\AppData\Local"

        set "RUNTIME_TMP=%LOCALAPPDATA%\{app_name}\_runtime_tmp"
        if not exist "%RUNTIME_TMP%" mkdir "%RUNTIME_TMP%" >nul 2>&1

        set "PYINSTALLER_RESET_ENVIRONMENT=1"
        set "_MEIPASS2="
        set "_PYI_APPLICATION_HOME_DIR="
        set "PYTHONHOME="
        set "PYTHONPATH="
        set "PYTHONNOUSERSITE="
        set "VIRTUAL_ENV="
        set "CONDA_PREFIX="
        set "__PYVENV_LAUNCHER__="

        set "PYINSTALLER_RUNTIME_TMPDIR=%RUNTIME_TMP%"
        set "TMP=%RUNTIME_TMP%"
        set "TEMP=%RUNTIME_TMP%"

        set "EXE={exe_name}"
        set "NEW={exe_name}.new"

        :waitproc
        tasklist | find /i "%EXE%" >nul
        if not errorlevel 1 (
            timeout /t 1 /nobreak >nul
            goto waitproc
        )

        for /l %%i in (1,1,60) do (
            if exist "%EXE%" del /f /q "%EXE%" >nul 2>&1
            if exist "%NEW%" ren "%NEW%" "%EXE%" >nul 2>&1
            if exist "%EXE%" if not exist "%NEW%" goto :run
            timeout /t 2 /nobreak >nul
        )

        goto :fail

        :run
        timeout /t 5 /nobreak >nul

        for /l %%j in (1,1,3) do (
            start "" "%EXE%"
            timeout /t 1 /nobreak >nul
        )

        exit /b 0

        :fail
        echo Update failed.
        pause
        exit /b 1
    """).strip()

    try:
        updater.write_text(cmd, encoding="utf-8")
    except Exception:
        pass

    try:
        if sys.platform.startswith("win"):
            import subprocess
            subprocess.Popen(['cmd.exe', '/c', str(updater)], cwd=str(app_folder), close_fds=True)
        else:
            webbrowser.open(github_releases_url)
    except Exception:
        webbrowser.open(github_releases_url)
