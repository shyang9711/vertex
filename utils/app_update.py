from __future__ import annotations

import os
import sys
import json
import re
import ssl
import urllib.request
import urllib.error
import webbrowser
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

# Optional logger (safe if you use your app_logging)
try:
    from functions.utils.app_logging import get_logger
except Exception:
    try:
        from utils.app_logging import get_logger
    except Exception:
        get_logger = None

LOG = get_logger("app_update") if get_logger else None


def _parse_version(s: str) -> tuple[int, int, int]:
    nums = re.findall(r"\d+", str(s))
    major = int(nums[0]) if len(nums) > 0 else 0
    minor = int(nums[1]) if len(nums) > 1 else 0
    patch = int(nums[2]) if len(nums) > 2 else 0
    return (major, minor, patch)


def _is_newer_version(latest: str, current: str) -> bool:
    return _parse_version(latest) > _parse_version(current)


def _fetch_latest_release_json_silent(github_api_latest: str, app_name: str) -> dict | None:
    try:
        req = urllib.request.Request(github_api_latest, headers={"User-Agent": app_name})
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, context=ctx, timeout=6) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def _fetch_update_policy_from_latest_release_silent(
    github_api_latest: str,
    app_name: str,
    policy_asset_name: str,
) -> dict | None:
    data = _fetch_latest_release_json_silent(github_api_latest, app_name)
    if not isinstance(data, dict):
        return None

    assets = data.get("assets") or []
    policy_url = None
    for a in assets:
        if a.get("name") == policy_asset_name:
            policy_url = a.get("browser_download_url")
            break

    if not policy_url:
        return None

    try:
        req2 = urllib.request.Request(policy_url, headers={"User-Agent": app_name})
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req2, context=ctx, timeout=6) as resp2:
            raw = resp2.read().decode("utf-8", errors="replace")
        obj = json.loads(raw)
        return obj if isinstance(obj, dict) else None
    except Exception:
        return None


def _latest_tag_from_release_json(release_json: dict) -> str | None:
    tag = str(release_json.get("tag_name") or "").strip()
    if tag.lower().startswith("v"):
        tag = tag[1:]
    return tag or None


def is_update_required_by_policy(
    app_version: str,
    github_api_latest: str,
    app_name: str,
    policy_asset_name: str = "update_policy.json",
) -> tuple[bool, str | None, str | None, dict | None]:
    """
    Returns (required, reason, latest_tag, policy_dict)
    - required = True only when policy.force == True and app_version < policy.min_required_version
    """
    release_json = _fetch_latest_release_json_silent(github_api_latest, app_name)
    latest_tag = _latest_tag_from_release_json(release_json) if isinstance(release_json, dict) else None

    policy = _fetch_update_policy_from_latest_release_silent(
        github_api_latest=github_api_latest,
        app_name=app_name,
        policy_asset_name=policy_asset_name,
    )

    if not isinstance(policy, dict):
        if LOG:
            LOG.warning(
                "update_policy.json not found or unreadable from latest release. "
                "latest_tag=%s policy_asset_name=%s",
                latest_tag, policy_asset_name
            )
        return (False, None, latest_tag, None)

    min_req = str(policy.get("min_required_version") or "").strip()
    force = bool(policy.get("force", False))
    msg = str(policy.get("message") or "Update required to continue.").strip()

    if force and min_req and _is_newer_version(min_req, app_version):
        reason = f"{msg}\n\nThis version ({app_version}) is below the minimum required version ({min_req})."
        return (True, reason, latest_tag, policy)

    return (False, None, latest_tag, policy)


def enforce_major_update_on_startup(
    root: tk.Tk,
    *,
    app_name: str,
    app_version: str,
    github_api_latest: str,
    github_releases_url: str,
    update_asset_name: str = "vertex.exe",
    policy_asset_name: str = "update_policy.json",
) -> bool:
    """
    If policy says update is required, show a blocking modal.
    Clicking Update calls check_for_updates().
    Returns True if blocked, else False.
    """
    required, reason, latest_tag, _policy = is_update_required_by_policy(
        app_version=app_version,
        github_api_latest=github_api_latest,
        app_name=app_name,
        policy_asset_name=policy_asset_name,
    )

    if LOG:
        LOG.info("policy=%s required=%s version=%s", _policy, required, app_version)

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
        f"Update required\n\n"
        f"Current version: {app_version}\n"
        f"Latest version:  {latest_tag or 'unknown'}\n\n"
        f"{reason}\n\n"
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
    Same behavior you currently have:
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
                "Once you create a release for this app, "
                "the update checker will compare versions.",
                parent=parent,
            )
            return
        messagebox.showerror("Update Check Failed", f"HTTP error from GitHub:\n{e}", parent=parent)
        return
    except Exception as e:
        messagebox.showerror("Update Check Failed", f"Could not contact GitHub:\n{e}", parent=parent)
        return

    tag = str(data.get("tag_name") or "").strip()
    if tag.lower().startswith("v"):
        tag = tag[1:]

    if not tag:
        messagebox.showinfo(
            "Updates",
            "Latest release has no tag_name.\nUse tags like v0.1.0 on GitHub releases.",
            parent=parent,
        )
        return

    if not _is_newer_version(tag, app_version):
        messagebox.showinfo(
            "Up to date",
            f"You are running version {app_version}, which is the latest release.",
            parent=parent,
        )
        return

    if not getattr(sys, "frozen", False):
        if messagebox.askyesno(
            "Update available",
            f"Current version: {app_version}\nLatest version: {tag}\n\nOpen the GitHub Releases page?",
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
            f"Current version: {app_version}\nLatest version: {tag}\n\n"
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

    # write & run updater cmd (kept consistent with your current behavior)
    import textwrap

    updater = app_folder / "update_vertex.cmd"

    cmd = textwrap.dedent(fr"""
        @echo off
        setlocal EnableExtensions
        echo Updating {app_name}...

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

        for /l %%j in (1,1,8) do (
            echo Starting %%j/8.
            start "" /d "%DIR%" "%DIR%\%EXE%"
            timeout /t 10 /nobreak >nul
            tasklist | find /i "%EXE%" >nul
            if not errorlevel 1 goto :cleanup
        )

        goto :fail

        :cleanup
        popd
        del "%~f0"
        exit /b 0

        :fail
        popd
        echo Update failed to relaunch.
        pause
        exit /b 1
    """).strip() + "\n"

    try:
        updater.write_text(cmd, encoding="utf-8")
        clean_env = os.environ.copy()
        for k in (
            "_MEIPASS2",
            "_PYI_APPLICATION_HOME_DIR",
            "PYTHONHOME",
            "PYTHONPATH",
            "PYTHONNOUSERSITE",
            "VIRTUAL_ENV",
            "CONDA_PREFIX",
            "__PYVENV_LAUNCHER__",
        ):
            clean_env.pop(k, None)

        subprocess.Popen(
            ["cmd.exe", "/d", "/c", "call", str(updater)],
            creationflags=subprocess.CREATE_NEW_CONSOLE,
            env=clean_env,
        )
    except Exception as e:
        messagebox.showerror("Update Failed", f"Updater error:\n{repr(e)}", parent=parent)
        return

    sys.exit(0)
