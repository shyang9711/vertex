from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import os, json, threading, subprocess, signal
from dataclasses import dataclass
from pathlib import Path
import shutil
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
try:
    from functions.utils.app_logging import get_logger
    from styles.new_ui import NewUI
except ModuleNotFoundError:
    from utils.app_logging import get_logger
    from styles.new_ui import NewUI
LOG = get_logger("actions")

@dataclass
class ToolSpec:
    key: str
    label: str
    script: str

def _tool_root() -> Path:
    """Return the directory where our tool scripts & scripts.json live."""
    # In PyInstaller onefile, --add-data files go under sys._MEIPASS
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    # Dev / normal Python: functions root (same as before)
    return Path(__file__).resolve().parent.parent

class ActionRunnerPage:
    """
    In-app 'cmd-like' runner for POS Parse, Validate Tax Payments, Vendor Match.
    - Uses current client context from app._current_client()
    - Tool names come from scripts.json (sibling to client_manager.py)
    - Streams child stdout/stderr live into console
    """
    # canonical filenames for the three integrated tools
    TARGET_SCRIPTS = ("pos_parse.py", "validate_tax_payments.py", "vendor_match.py", "de9c_to_csv.py", "file_name_change.py")

    def __init__(self, app):
        self.app = app
        self.log = get_logger("actions")
        self.frame: ttk.Frame | None = None
        self._proc: subprocess.Popen | None = None
        self._reader_thread: threading.Thread | None = None
        self._stop_flag = False
        self._current_tool: ToolSpec | None = None
        self._tools: list[ToolSpec] = []
        self._load_tools_from_scripts_json()

    # --------- Public API ---------
    def ensure(self, parent: tk.Widget) -> ttk.Frame:
        self.log.info("ensure() building action runner UI")
        if self.frame is not None and self.frame.winfo_exists():
            return self.frame

        root = self.frame = ttk.Frame(parent, style="Card.TFrame")
        root.pack_forget()

        # Header
        header = ttk.Frame(root, padding=(12, 10)); header.pack(fill="x")
        ttk.Label(header, text="Client Actions", style="Title.TLabel").pack(side="left")

        # Controls (no client selector)
        ctrls = ttk.Frame(root, padding=(12, 6)); ctrls.pack(fill="x")

        ttk.Label(ctrls, text="Tool:").pack(side="left")
        labels = [t.label for t in self._tools]
        self.var_tool = tk.StringVar(value=labels[0] if labels else "")
        self.dd_tool = ttk.Combobox(
            ctrls,
            textvariable=self.var_tool,
            width=36,
            state="readonly",
            values=labels
        )
        self.dd_tool.pack(side="left", padx=(6, 14))

        self.btn_run   = ttk.Button(ctrls, text="Run",   command=self._on_run,  style="NewUI.TButton", width=10)
        self.btn_stop  = ttk.Button(ctrls, text="Stop",  command=self._on_stop, style="NewUI.TButton", width=10, state="disabled")
        self.btn_clear = ttk.Button(ctrls, text="Clear", command=self._clear_console,                  width=10)
        self.btn_run.pack(side="left"); self.btn_stop.pack(side="left", padx=(6,0)); self.btn_clear.pack(side="left", padx=(6,0))

        # Console
        wrap = ttk.Frame(root, padding=(12, 6)); wrap.pack(fill="both", expand=True)
        self.console = ScrolledText(wrap, height=26, wrap="word")
        self.console.pack(fill="both", expand=True)
        self._log("Ready.\n")

        # Footer
        self.status = tk.StringVar(value="")
        ttk.Label(root, textvariable=self.status, anchor="w").pack(fill="x", padx=12, pady=(0,10))

        return root

    def preselect(self, tool_key: str | None):
        """Optionally preselect a tool when navigating here."""
        if tool_key and any(t.key == tool_key for t in self._tools):
            self.var_tool.set(tool_key)

    # --------- Internals ---------
    def _load_tools_from_scripts_json(self):
        """
        Read friendly names from scripts.json at <functions>/scripts.json and
        build ToolSpec list for our TARGET_SCRIPTS.
        """
        self.log.info("Loading scripts.json for tool labels")
        try:
            # In onefile builds, bundled resources live under sys._MEIPASS
            if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
                functions_dir = Path(sys._MEIPASS)
            else:
                functions_dir = _tool_root()
            scripts_json = functions_dir / "scripts.json"
            mapping = {}
            if scripts_json.exists():
                with open(scripts_json, "r", encoding="utf-8") as f:
                    data = json.load(f) or {}
                for fname in self.TARGET_SCRIPTS:
                    meta = data.get(fname, {})
                    label = str(meta.get("name") or Path(fname).stem)
                    key = Path(fname).stem
                    
                    if (functions_dir / fname).exists():
                        mapping[key] = ToolSpec(key=key, label=label, script=fname)
            else:
                for fname in self.TARGET_SCRIPTS:
                    key = Path(fname).stem
                    mapping[key] = ToolSpec(key=key, label=key.replace("_", " ").title(), script=fname)

            self._tools = list(mapping.values())
            order = {Path(s).stem: i for i, s in enumerate(self.TARGET_SCRIPTS)}
            self._tools.sort(key=lambda t: order.get(t.key, 999))
        except Exception as e:
            LOG.exception("Failed to load tools list: %s", e)
            self._tools = [ToolSpec(Path(s).stem, Path(s).stem.replace("_"," ").title(), s) for s in self.TARGET_SCRIPTS]

    def _resolve_tool(self) -> ToolSpec | None:
        label = (self.var_tool.get() or "").strip()
        for t in self._tools:
            if t.label == label:
                return t
        return self._tools[0] if self._tools else None

    def _on_run(self):
        self.log.info("Run clicked; tool=%s", self.var_tool.get())
        if self._proc and self._proc.poll() is None:
            messagebox.showinfo("Actions", "A tool is already running. Stop it first.")
            return

        tool = self._resolve_tool()
        if not tool:
            messagebox.showerror("Actions", "No tool available.")
            return

        env = os.environ.copy()

        # Only scrub PyInstaller-related vars if we're launching an external interpreter.
        # In frozen mode we are spawning THIS exe, so keep them intact.
        if not getattr(sys, "frozen", False):
            for k in list(env.keys()):
                if k.startswith("_PYI_") or k.startswith("PYI_") or k.startswith("_MEI"):
                    env.pop(k, None)
            for k in ("PYTHONHOME", "PYTHONPATH"):
                env.pop(k, None)

        env["PYTHONIOENCODING"] = "utf-8"

        functions_dir = _tool_root()
        tool_path = functions_dir / tool.script
        if not tool_path.exists():
            messagebox.showerror("Actions", f"Tool not found: {tool_path}")
            return

        self._clear_console()
        self._log(f"Running {tool.label} → {tool_path}\n\n")
        self.status.set(f"Running {tool.label}...")
        self._current_tool = tool
        self._stop_flag = False

        # Choose interpreter:
        # - Dev: use the current Python
        # - Frozen: try system Python (python / python3)
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--run-tool", tool.script]
        else:
            cmd = [sys.executable, str(tool_path)]

        work_dir = Path(tempfile.gettempdir()) / "VertexToolWork"
        try:
            work_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            work_dir = Path.home()

        try:
            if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
                mei_root = Path(sys._MEIPASS)
                if mei_root in tool_path.resolve().parents:
                    tool_copy = work_dir / tool_path.name
                    shutil.copy2(tool_path, tool_copy)
                    tool_path = tool_copy  # run the copied script instead
        except Exception:
            pass

        _path_parts = [p for p in (env.get("PATH", "").split(os.pathsep)) if "_MEI" not in p and "_PYI" not in p]
        env["PATH"] = os.pathsep.join(_path_parts)

        try:
            self._proc = subprocess.Popen(
                cmd,
                cwd=str(work_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env,
            )
        except Exception as e:
            LOG.exception("Subprocess launch failed for %s: %s", tool.script, e)
            self._log(f"[ERROR] Failed to start: {e}\n")
            self.status.set("Failed to start.")
            return


        # Stream output in a background thread
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")

        def _reader():
            try:
                for line in self._proc.stdout:
                    if self._stop_flag:
                        break
                    self._log(line)
            except Exception as e:
                self._log(f"[ReaderError] {e}\n")
            finally:
                rc = self._proc.poll()
                if rc is None:
                    try:
                        self._proc.wait(timeout=0.2)
                        rc = self._proc.returncode
                    except Exception:
                        rc = None
                self._proc = None
                self._log(f"\n[Process exited] returncode={rc}\n")
                self.app.after(0, lambda: (self.btn_run.config(state="normal"), self.btn_stop.config(state="disabled"),
                                           self.status.set("Done." if rc == 0 else f"Exited ({rc})")))
        self._reader_thread = threading.Thread(target=_reader, daemon=True)
        self._reader_thread.start()

    def _on_stop(self):
        self.log.info("Stop clicked (proc=%s)", bool(self._proc))
        if not self._proc:
            return
        self._stop_flag = True
        try:
            if os.name == "nt":
                self._proc.send_signal(signal.CTRL_BREAK_EVENT)
            else:
                self._proc.terminate()
        except Exception:
            try:
                self._proc.kill()
            except Exception:
                pass
        self._log("\n[Requested stop]\n")
        self.status.set("Stopping...")

    def _clear_console(self):
        self.console.delete("1.0", "end")

    def _log(self, text: str):
        self.console.insert("end", text)
        self.console.see("end")
