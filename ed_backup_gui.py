import json
import os
import platform
import subprocess
import threading
import queue
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import zipfile
from typing import List, Tuple, Optional

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image  # pip install pillow


# -----------------------------
# THEME COLORS (Elite orange)
# -----------------------------
ELITE_ORANGE = ("#ff7a00", "#ff7a00")      # progress bar fill (light, dark)
TRACK_DARK = ("#2a2a2a", "#2a2a2a")        # progress bar track/background
ELITE_ORANGE_TEXT = "#ff7a00"             # label text color (single hex)


@dataclass
class BackupSelection:
    journal: bool
    bindings: bool
    graphics: bool


def default_paths_windows() -> Tuple[Path, Path, Path]:
    """Return default Elite Dangerous paths on Windows."""
    user_home = Path.home()

    journal = user_home / "Saved Games" / "Frontier Developments" / "Elite Dangerous"

    local_appdata = Path(os.environ.get("LOCALAPPDATA", str(user_home / "AppData" / "Local")))
    bindings = local_appdata / "Frontier Developments" / "Elite Dangerous" / "Options" / "Bindings"
    graphics = local_appdata / "Frontier Developments" / "Elite Dangerous" / "Options" / "Graphics"

    return journal, bindings, graphics


def iter_files(folder: Path):
    """Yield all files under folder recursively."""
    if not folder.exists():
        return
    for p in folder.rglob("*"):
        if p.is_file():
            yield p


def open_folder_in_explorer(folder: Path) -> bool:
    """Open a folder in system file explorer. Returns True if successful."""
    try:
        folder = folder.expanduser().resolve()
        if not folder.exists():
            return False

        system = platform.system().lower()
        if system == "windows":
            os.startfile(str(folder))  # type: ignore[attr-defined]
            return True
        elif system == "darwin":
            subprocess.Popen(["open", str(folder)])
            return True
        else:
            subprocess.Popen(["xdg-open", str(folder)])
            return True
    except Exception:
        return False


def is_windows_process_running(process_names: List[str]) -> Tuple[bool, List[str]]:
    """
    Windows-only: check if any of the given process names are running.
    Returns (is_running, matched_names).
    """
    if platform.system().lower() != "windows":
        return False, []

    try:
        out = subprocess.check_output(["tasklist"], text=True, errors="ignore")
        out_lower = out.lower()

        matched = []
        for name in process_names:
            if name.lower() in out_lower:
                matched.append(name)

        return (len(matched) > 0), matched
    except Exception:
        return False, []


class EDBackupApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Elite Dangerous – Backup Tool (Journal / Bindings / Graphics)")
        self.geometry("940x720")
        self.minsize(940, 720)

        # Lock window size so the banner layout stays consistent
        self.resizable(False, False)

        # Window icon
        self._set_window_icon()

        self.journal_path, self.bindings_path, self.graphics_path = self._detect_paths()

        # Checkboxes
        self.var_journal = ctk.BooleanVar(value=True)
        self.var_bindings = ctk.BooleanVar(value=True)
        self.var_graphics = ctk.BooleanVar(value=True)

        # Source paths
        self.src_journal = ctk.StringVar(value=str(self.journal_path))
        self.src_bindings = ctk.StringVar(value=str(self.bindings_path))
        self.src_graphics = ctk.StringVar(value=str(self.graphics_path))

        # Destination folder
        self.dest_folder = ctk.StringVar(value=str(Path.home() / "Desktop"))

        # Status + progress
        self.status = ctk.StringVar(value="Ready.")
        self.progress_text = ctk.StringVar(value="")

        self._progress_queue: "queue.Queue[dict]" = queue.Queue()
        self._worker_thread: Optional[threading.Thread] = None
        self._is_working = False

        self._ui()
        self._poll_progress_queue()

    # -----------------------------
    # Helpers
    # -----------------------------
    def _script_dir(self) -> Path:
        """Return directory where this script is located."""
        try:
            return Path(__file__).resolve().parent
        except NameError:
            return Path.cwd()

    def _set_window_icon(self):
        """Set window icon. Prefer .ico on Windows, fallback to .png."""
        script_dir = self._script_dir()
        ico_path = script_dir / "ed_icon.ico"
        png_path = script_dir / "ed_icon.png"

        try:
            if ico_path.exists():
                self.iconbitmap(str(ico_path))
                return
        except Exception:
            pass

        try:
            if png_path.exists():
                img = tk.PhotoImage(file=str(png_path))
                self.iconphoto(True, img)
                return
        except Exception:
            pass

    def _detect_paths(self) -> Tuple[Path, Path, Path]:
        """Detect default paths (Windows expected)."""
        if platform.system().lower() == "windows":
            return default_paths_windows()
        home = Path.home()
        return home, home, home

    def _warn_if_game_running(self) -> bool:
        """Warn if Elite Dangerous / EDMC appears to be running (Windows)."""
        watch = [
            "EliteDangerous64.exe",
            "EliteDangerous32.exe",
            "EliteDangerous.exe",
            "EDLaunch.exe",
            "EDMarketConnector.exe",
        ]
        running, matched = is_windows_process_running(watch)
        if not running:
            return True

        msg = (
            "It looks like Elite Dangerous and/or EDMC might be running:\n\n"
            f"{', '.join(matched)}\n\n"
            "Recommended: close the game (and EDMC) before making a backup.\n\n"
            "Do you want to continue anyway?"
        )
        return messagebox.askyesno("Warning", msg)

    # -----------------------------
    # UI
    # -----------------------------
    def _ui(self):
        header = ctk.CTkFrame(self, corner_radius=18)
        header.pack(fill="x", padx=16, pady=(16, 10))

        banner_path = self._script_dir() / "ed_banner.png"
        if banner_path.exists():
            banner_img = Image.open(banner_path).convert("RGBA")
            self._banner_ctk = ctk.CTkImage(light_image=banner_img, dark_image=banner_img, size=(940, 160))
            ctk.CTkLabel(header, image=self._banner_ctk, text="").pack(padx=0, pady=0)
        else:
            ctk.CTkLabel(header, text="Elite Dangerous Backup", font=ctk.CTkFont(size=24, weight="bold")).pack(
                anchor="w", padx=16, pady=(14, 0)
            )
            ctk.CTkLabel(
                header,
                text="Each selected category is exported into a separate ZIP file.",
                font=ctk.CTkFont(size=13),
                text_color="#b0b0b0",
            ).pack(anchor="w", padx=16, pady=(0, 14))

        body = ctk.CTkFrame(self, corner_radius=18)
        body.pack(fill="both", expand=True, padx=16, pady=10)

        # Sources
        sources = ctk.CTkFrame(body, corner_radius=14)
        sources.pack(fill="x", padx=14, pady=14)
        ctk.CTkLabel(sources, text="What to back up", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(10, 8)
        )

        self._row_source(sources, 1, "Journal (game logs)", self.var_journal, self.src_journal)
        self._row_source(sources, 2, "Bindings (key mappings)", self.var_bindings, self.src_bindings)
        self._row_source(sources, 3, "Graphics (settings)", self.var_graphics, self.src_graphics)

        # Destination
        dest = ctk.CTkFrame(body, corner_radius=14)
        dest.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkLabel(dest, text="Save ZIP files to", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 8)
        )

        ctk.CTkEntry(dest, textvariable=self.dest_folder, height=36).grid(
            row=1, column=0, sticky="ew", padx=12, pady=(0, 12)
        )
        ctk.CTkButton(dest, text="Choose folder…", width=160, command=self.pick_dest).grid(
            row=1, column=1, padx=(0, 12), pady=(0, 12)
        )
        dest.grid_columnconfigure(0, weight=1)

        # Actions
        actions = ctk.CTkFrame(body, corner_radius=14)
        actions.pack(fill="x", padx=14, pady=(0, 14))

        self.btn_backup = ctk.CTkButton(actions, text="Create backup ZIP(s)", height=42, command=self.make_backup)
        self.btn_backup.pack(side="left", padx=12, pady=12)

        ctk.CTkButton(actions, text="How to restore?", height=42, command=self.restore_hint).pack(
            side="left", padx=12, pady=12
        )
        ctk.CTkButton(actions, text="Open destination folder", height=42, command=self.open_dest_folder).pack(
            side="left", padx=12, pady=12
        )

        # Progress
        progress_frame = ctk.CTkFrame(body, corner_radius=14)
        progress_frame.pack(fill="x", padx=14, pady=(0, 14))

        ctk.CTkLabel(progress_frame, text="Progress", font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            height=18,
            progress_color=ELITE_ORANGE,
            fg_color=TRACK_DARK,
            mode="determinate",  # IMPORTANT
        )
        self.progress_bar.pack(fill="x", padx=12, pady=(0, 6))
        self.progress_bar.set(0.0)

        self.progress_label = ctk.CTkLabel(
            progress_frame,
            textvariable=self.progress_text,
            font=ctk.CTkFont(size=12),
            text_color=ELITE_ORANGE_TEXT,
        )
        self.progress_label.pack(anchor="w", padx=12, pady=(0, 10))

        # Status bar
        status_frame = ctk.CTkFrame(self, corner_radius=14)
        status_frame.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkLabel(status_frame, textvariable=self.status, font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=12, pady=10
        )

    def _row_source(self, parent, row: int, label: str, var_check, var_path):
        chk = ctk.CTkCheckBox(parent, text=label, variable=var_check)
        chk.grid(row=row, column=0, sticky="w", padx=12, pady=6)

        ctk.CTkEntry(parent, textvariable=var_path, height=34).grid(
            row=row, column=1, sticky="ew", padx=(10, 10), pady=6
        )

        def browse():
            p = filedialog.askdirectory(title=f"Select folder: {label}")
            if p:
                var_path.set(p)

        ctk.CTkButton(parent, text="Browse…", width=110, command=browse).grid(
            row=row, column=2, padx=(0, 12), pady=6
        )

        parent.grid_columnconfigure(1, weight=1)

    # -----------------------------
    # Buttons
    # -----------------------------
    def pick_dest(self):
        p = filedialog.askdirectory(title="Select destination folder for ZIP files")
        if p:
            self.dest_folder.set(p)

    def open_dest_folder(self):
        dest = Path(self.dest_folder.get()).expanduser()
        if not open_folder_in_explorer(dest):
            messagebox.showerror("Error", "Could not open destination folder. Check the path.")

    # -----------------------------
    # Progress polling
    # -----------------------------
    def _poll_progress_queue(self):
        try:
            while True:
                msg = self._progress_queue.get_nowait()
                kind = msg.get("type")

                if kind == "start":
                    self._is_working = True
                    self.btn_backup.configure(state="disabled")
                    self.progress_bar.set(0.0)
                    self.progress_text.set("Starting…")
                    self.status.set("Working…")

                elif kind == "progress":
                    total = max(1, int(msg.get("total", 1)))
                    done = max(0, int(msg.get("done", 0)))
                    current = msg.get("current", "")

                    frac = done / total
                    if frac < 0:
                        frac = 0.0
                    if frac > 1:
                        frac = 1.0

                    self.progress_bar.set(frac)
                    self.progress_text.set(current or f"{done}/{total}")

                elif kind == "done":
                    self._is_working = False
                    self.btn_backup.configure(state="normal")
                    self.progress_bar.set(1.0)
                    self.progress_text.set("Done.")
                    self.status.set(msg.get("status", "Done."))

                    summary = msg.get("summary", "Backup complete.")
                    messagebox.showinfo("Backup complete", summary)

                elif kind == "error":
                    self._is_working = False
                    self.btn_backup.configure(state="normal")
                    self.progress_bar.set(0.0)
                    self.progress_text.set("")
                    self.status.set("Error.")
                    messagebox.showerror("Error", msg.get("error", "Unknown error."))

        except queue.Empty:
            pass

        self.after(80, self._poll_progress_queue)  # a bit faster updates

    # -----------------------------
    # Worker: reliable progress
    # -----------------------------
    def _worker_backup(self, tasks: List[Tuple[str, Path, str]], dest: Path, timestamp: str):
        """
        Reliable progress:
        1) build file lists first -> known total
        2) zip and update progress every N files
        """
        try:
            existing: List[Tuple[str, Path, str, List[Path]]] = []
            skipped: List[str] = []

            for name, folder, arc_root in tasks:
                if not folder.exists():
                    skipped.append(f"{name} (missing folder: {folder})")
                    continue

                files = list(iter_files(folder))
                existing.append((name, folder, arc_root, files))

            if not existing:
                self._progress_queue.put({"type": "error", "error": "No ZIP files created. Check your source folder paths."})
                return

            total_files = sum(len(files) for _, _, _, files in existing)
            total_units = total_files + len(existing)  # + manifest.json for each zip
            if total_units <= 0:
                total_units = 1

            done_units = 0
            created: List[Tuple[Path, int]] = []

            # Update every N files to keep UI smooth
            update_every = 15
            last_push = 0.0
            min_interval = 0.05  # seconds

            for name, folder, arc_root, files in existing:
                zip_path = dest / f"EliteDangerous_{name}_Backup_{timestamp}.zip"
                files_in_zip = 0

                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    # manifest
                    manifest = {
                        "created_at": datetime.now().isoformat(timespec="seconds"),
                        "machine": platform.node(),
                        "os": platform.platform(),
                        "backup_type": name,
                        "source_folder": str(folder),
                    }
                    zf.writestr("manifest.json", json.dumps(manifest, indent=2))
                    done_units += 1
                    self._progress_queue.put({
                        "type": "progress",
                        "total": total_units,
                        "done": done_units,
                        "current": f"{name}: writing manifest.json"
                    })

                    # files
                    for idx, p in enumerate(files, start=1):
                        arcname = str(Path(arc_root) / p.relative_to(folder))
                        zf.write(p, arcname=arcname)
                        files_in_zip += 1
                        done_units += 1

                        # throttle UI updates
                        now = time.time()
                        if (idx % update_every == 0) or (now - last_push >= min_interval) or (idx == len(files)):
                            short = str(p)
                            if len(short) > 95:
                                short = "…" + short[-94:]
                            self._progress_queue.put({
                                "type": "progress",
                                "total": total_units,
                                "done": done_units,
                                "current": f"{name}: {short}"
                            })
                            last_push = now

                created.append((zip_path, files_in_zip))

            lines = []
            for zp, n in created:
                lines.append(f"✅ {zp.name}  (files: {n})")
            if skipped:
                lines.append("\nSkipped:")
                for s in skipped:
                    lines.append(f"⚠️ {s}")

            self._progress_queue.put({
                "type": "done",
                "status": f"Done: created {len(created)} ZIP file(s).",
                "summary": "\n".join(lines)
            })

        except Exception as e:
            self._progress_queue.put({"type": "error", "error": f"Failed to create ZIP file(s):\n{e}"})

    # -----------------------------
    # Main action
    # -----------------------------
    def make_backup(self):
        if self._is_working:
            messagebox.showinfo("Busy", "Backup is already running.")
            return

        sel = BackupSelection(
            journal=self.var_journal.get(),
            bindings=self.var_bindings.get(),
            graphics=self.var_graphics.get(),
        )

        if not (sel.journal or sel.bindings or sel.graphics):
            messagebox.showerror("Error", "Select at least one category to back up.")
            return

        dest = Path(self.dest_folder.get()).expanduser()
        if not dest.exists():
            messagebox.showerror("Error", "Destination folder does not exist.")
            return

        if not self._warn_if_game_running():
            self.status.set("Cancelled.")
            return

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        journal = Path(self.src_journal.get()).expanduser()
        bindings = Path(self.src_bindings.get()).expanduser()
        graphics = Path(self.src_graphics.get()).expanduser()

        tasks: List[Tuple[str, Path, str]] = []
        if sel.journal:
            tasks.append(("Journal", journal, "journal"))
        if sel.bindings:
            tasks.append(("Bindings", bindings, "bindings"))
        if sel.graphics:
            tasks.append(("Graphics", graphics, "graphics"))

        self._progress_queue.put({"type": "start"})
        self._worker_thread = threading.Thread(
            target=self._worker_backup,
            args=(tasks, dest, timestamp),
            daemon=True
        )
        self._worker_thread.start()

    def restore_hint(self):
        msg = (
            "How to restore:\n\n"
            "1) Close Elite Dangerous (and EDMC, if you use it).\n"
            "2) Extract the relevant ZIP file.\n"
            "3) Copy the extracted folder contents back to the correct location:\n"
            "   - Journal ZIP:   journal -> Saved Games\\Frontier Developments\\Elite Dangerous\n"
            "   - Bindings ZIP:  bindings -> AppData\\Local\\Frontier Developments\\Elite Dangerous\\Options\\Bindings\n"
            "   - Graphics ZIP:  graphics  -> AppData\\Local\\Frontier Developments\\Elite Dangerous\\Options\\Graphics\n\n"
            "Tip: If the folders already exist, overwrite files (or make a backup first)."
        )
        messagebox.showinfo("Restore instructions", msg)


if __name__ == "__main__":
    app = EDBackupApp()
    app.mainloop()
