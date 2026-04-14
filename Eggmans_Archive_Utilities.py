# -*- coding: utf-8 -*-
"""
Eggman's Archive Utilities  v2.0
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Tab 1 – Extractor   : Recursive ZIP / 7Z / RAR extraction
         • Same-location or mirrored destination
         • Keep / Recycle Bin / Permanent delete after extract
         • Move archives (mirrored or flat) after extract
         • Nested archive detection + auto-extraction
         • Single-file-in-zip optimization preserved
         • Double-nesting flattener preserved
Tab 2 – ZIP Store Packer : Wrap files in ZIP_STORED containers
         • Multi-extension targeting with pill tags
         • Verify-before-delete
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Dependencies (all optional extras):
    pip install tkinterdnd2    # drag-and-drop
    pip install send2trash     # recycle-bin delete
"""

import re
import sys
import shutil
import zipfile
import threading
import subprocess
import tkinter as tk
import time
from collections import deque
from tkinter import ttk, filedialog, scrolledtext
from pathlib import Path
from datetime import timedelta

# ── Optional dependencies ──────────────────────────────────────────────────────

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

try:
    import send2trash
    TRASH_AVAILABLE = True
except ImportError:
    TRASH_AVAILABLE = False

# ── Constants ──────────────────────────────────────────────────────────────────

SEVEN_ZIP       = r"C:\Program Files\7-Zip-Zstandard\7z.exe"
INVALID_WIN_CHARS = r'<>:"/\|?*'

# ── Palette ────────────────────────────────────────────────────────────────────

C = {
    "bg":      "#0d0f1a",
    "bg2":     "#161929",
    "bg3":     "#1f2240",
    "bg4":     "#252a4a",
    "accent":  "#7c6ff7",
    "cyan":    "#38bdf8",
    "green":   "#4ade80",
    "amber":   "#fbbf24",
    "red":     "#f87171",
    "text":    "#dde3f0",
    "muted":   "#5a6382",
    "border":  "#2c3260",
}

# ══════════════════════════════════════════════════════════════════════════════
#  CORE LIBRARY
# ══════════════════════════════════════════════════════════════════════════════

def ensure_7z():
    if not Path(SEVEN_ZIP).is_file():
        raise FileNotFoundError(f"7z.exe not found at:\n  {SEVEN_ZIP}")


def run_7z(args):
    p = subprocess.run(
        [SEVEN_ZIP] + args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        errors="replace",
    )
    return p.returncode, p.stdout, p.stderr


def sanitize(name: str) -> str:
    name = re.sub(f"[{re.escape(INVALID_WIN_CHARS)}]", "_", name)
    name = name.rstrip(" .")
    return name or "extracted"


# ── Archive classification ─────────────────────────────────────────────────────

def _classify_zip_native(path: Path):
    """Fast Python classification for .zip — preserves original logic."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            infos = zf.infolist()
    except Exception:
        return "bad", None

    files, has_dir = [], False
    for info in infos:
        n = info.filename
        if n.endswith("/"):
            has_dir = True
            continue
        if "/" in n:
            has_dir = True
        files.append(n)

    if len(files) == 1 and not has_dir and "/" not in files[0]:
        return "single", files[0]
    return "folder", None


def _classify_via_7z(path: Path):
    """
    For 7z/rar: just verify the archive is readable.
    Always extract to its own folder — skip the single-file optimisation
    for non-zip formats to avoid silent misclassification from -slt parsing.
    """
    rc, _, _ = run_7z(["l", str(path)])
    if rc != 0:
        return "bad", None
    return "folder", None


def classify(path: Path):
    ext = path.suffix.lower()
    return _classify_zip_native(path) if ext == ".zip" else _classify_via_7z(path)


# ── Extraction helpers ─────────────────────────────────────────────────────────

def _merge_dir(src: Path, dst: Path):
    """Recursively merge src into dst."""
    for item in src.iterdir():
        d = dst / item.name
        if d.exists():
            if d.is_dir() and item.is_dir():
                _merge_dir(item, d)
                shutil.rmtree(item, ignore_errors=True)
            else:
                if d.is_dir():
                    shutil.rmtree(d, ignore_errors=True)
                else:
                    d.unlink(missing_ok=True)
                shutil.move(str(item), str(d))
        else:
            shutil.move(str(item), str(d))


def _flatten_double_nest(target: Path):
    """
    If target/ contains exactly one directory whose name == target.name,
    promote its contents up one level (eliminates tool-created double-nesting).
    """
    children = list(target.iterdir())
    if len(children) != 1:
        return
    only = children[0]
    if not only.is_dir() or only.name != target.name:
        return
    _merge_dir(only, target)
    shutil.rmtree(only, ignore_errors=True)


def extract_single(archive: Path, out_dir: Path):
    """Extract a single-file archive flat into out_dir."""
    out_dir.mkdir(parents=True, exist_ok=True)
    rc, _, err = run_7z(["e", "-y", f"-o{out_dir}", str(archive)])
    return rc == 0, err.strip()


def extract_to_folder(archive: Path, target: Path):
    """Extract multi-file archive into target/, with double-nest fix."""
    target.mkdir(parents=True, exist_ok=True)
    tmp = target / "__tmp_extract__"
    if tmp.exists():
        shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True, exist_ok=True)

    rc, _, err = run_7z(["x", "-y", f"-o{tmp}", str(archive)])
    if rc != 0:
        shutil.rmtree(tmp, ignore_errors=True)
        return False, err.strip()

    _merge_dir(tmp, target)
    shutil.rmtree(tmp, ignore_errors=True)
    _flatten_double_nest(target)
    return True, ""


def delete_archive(path: Path, mode: str):
    """mode: 'keep' | 'recycle' | 'permanent'"""
    if mode == "keep":
        return True, ""
    if mode == "recycle":
        if not TRASH_AVAILABLE:
            return False, "send2trash not installed"
        try:
            send2trash.send2trash(str(path))
            return True, ""
        except Exception as e:
            return False, str(e)
    try:
        path.unlink()
        return True, ""
    except Exception as e:
        return False, str(e)


def move_mirrored(archive: Path, src_root: Path, move_root: Path):
    """Move archive to move_root, preserving its path relative to src_root."""
    try:
        rel  = archive.relative_to(src_root)
        dest = move_root / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(archive), str(dest))
        return True, str(dest)
    except Exception as e:
        return False, str(e)


def move_flat(archive: Path, move_root: Path):
    """Move archive into move_root with no subdirectory structure.
    Collisions are renamed  name(1).ext, name(2).ext …"""
    try:
        move_root.mkdir(parents=True, exist_ok=True)
        dest = move_root / archive.name
        if dest.exists():
            stem, suffix = archive.stem, archive.suffix
            n = 1
            while dest.exists():
                dest = move_root / f"{stem}({n}){suffix}"
                n += 1
        shutil.move(str(archive), str(dest))
        return True, str(dest)
    except Exception as e:
        return False, str(e)


def scan_for_archives(folder: Path, exts: set) -> list[Path]:
    """Return all archive files found recursively inside folder."""
    found = []
    for ext in exts:
        found.extend(folder.rglob(f"*{ext}"))
    return sorted(found)


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED WIDGETS
# ══════════════════════════════════════════════════════════════════════════════

class FolderPicker(tk.Frame):
    """Drop zone + entry + browse button, reusable."""

    def __init__(self, parent, label="Drop folder here  —  or  Browse →",
                 disabled=False, **kwargs):
        super().__init__(parent, bg=C["bg2"], **kwargs)
        self._disabled = disabled
        self._var = tk.StringVar()
        self._callbacks = []

        state = "disabled" if disabled else "normal"
        zone_fg = C["muted"]
        zone_bg = C["bg3"] if not disabled else C["bg2"]

        self.zone = tk.Label(
            self, text=label,
            bg=zone_bg, fg=zone_fg,
            font=("Segoe UI", 9),
            pady=8, cursor="hand2" if not disabled else "",
            relief="flat"
        )
        self.zone.pack(fill="x", padx=0, pady=(0, 4))
        if not disabled:
            self.zone.bind("<Button-1>", lambda e: self._browse())

        row = tk.Frame(self, bg=C["bg2"])
        row.pack(fill="x")

        self.entry = tk.Entry(
            row, textvariable=self._var,
            bg=C["bg3"] if not disabled else C["bg"],
            fg=C["text"], insertbackground=C["text"],
            disabledbackground=C["bg"], disabledforeground=C["muted"],
            relief="flat", font=("Segoe UI", 9),
            state=state
        )
        self.entry.pack(side="left", fill="x", expand=True)

        self.btn = tk.Button(
            row, text="Browse",
            bg=C["accent"] if not disabled else C["bg3"],
            fg="white" if not disabled else C["muted"],
            relief="flat", font=("Segoe UI", 9, "bold"),
            command=self._browse, cursor="hand2" if not disabled else "",
            padx=10, state=state
        )
        self.btn.pack(side="right", padx=(6, 0))

        if DND_AVAILABLE and not disabled:
            self.zone.drop_target_register(DND_FILES)
            self.zone.dnd_bind("<<Drop>>", self._on_drop)
            self.entry.drop_target_register(DND_FILES)
            self.entry.dnd_bind("<<Drop>>", self._on_drop)

    def _browse(self):
        d = filedialog.askdirectory()
        if d:
            self.set(d)

    def _on_drop(self, event):
        self.set(event.data.strip().strip("{}"))

    def set(self, path: str):
        self._var.set(path)
        name = Path(path).name if path else ""
        self.zone.config(text=f"📂  {name}" if name else "Drop folder here  —  or  Browse →",
                         fg=C["cyan"] if name else C["muted"])
        for cb in self._callbacks:
            cb(path)

    def enable(self, yes: bool):
        state = "normal" if yes else "disabled"
        self.entry.config(state=state,
                          bg=C["bg3"] if yes else C["bg"],
                          fg=C["text"] if yes else C["muted"])
        self.btn.config(state=state,
                        bg=C["accent"] if yes else C["bg3"],
                        fg="white" if yes else C["muted"],
                        cursor="hand2" if yes else "")
        self.zone.config(cursor="hand2" if yes else "",
                         bg=C["bg3"] if yes else C["bg2"])
        if yes:
            self.zone.bind("<Button-1>", lambda e: self._browse())
        else:
            self.zone.unbind("<Button-1>")
        if DND_AVAILABLE:
            if yes:
                self.zone.drop_target_register(DND_FILES)
                self.zone.dnd_bind("<<Drop>>", self._on_drop)
                self.entry.drop_target_register(DND_FILES)
                self.entry.dnd_bind("<<Drop>>", self._on_drop)

    def get(self) -> str:
        return self._var.get().strip()

    def on_change(self, cb):
        self._callbacks.append(cb)


class LogPane(tk.Frame):
    TAGS = {
        "ok":     C["green"],
        "fail":   C["red"],
        "warn":   C["amber"],
        "info":   C["cyan"],
        "mute":   C["muted"],
        "skip":   C["muted"],
        "nested": C["accent"],
    }

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C["bg"], **kwargs)
        self.text = scrolledtext.ScrolledText(
            self, bg=C["bg2"], fg=C["text"],
            font=("Consolas", 9), relief="flat",
            insertbackground=C["text"], state="disabled",
            wrap="none", height=6
        )
        self.text.pack(fill="both", expand=True)

        hbar = tk.Scrollbar(self, orient="horizontal",
                            command=self.text.xview,
                            bg=C["bg3"], troughcolor=C["bg2"],
                            activebackground=C["bg4"])
        hbar.pack(fill="x")
        self.text.config(xscrollcommand=hbar.set)

        for tag, color in self.TAGS.items():
            self.text.tag_configure(tag, foreground=color)
        self.text.tag_configure("nested", foreground=C["accent"],
                                font=("Consolas", 9, "bold"))

    def write(self, tag: str, msg: str):
        def _do():
            self.text.config(state="normal")
            self.text.insert("end", msg, tag)
            self.text.see("end")
            self.text.config(state="disabled")
        self.after(0, _do)

    def clear(self):
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.config(state="disabled")

    def save(self):
        content = self.text.get("1.0", "end").strip()
        if not content:
            return
        p = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="archive_utility_log.txt"
        )
        if p:
            Path(p).write_text(content, encoding="utf-8")


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — EXTRACTOR
# ══════════════════════════════════════════════════════════════════════════════

class ExtractorTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=C["bg"])
        self._stop  = False
        self._build()

    def _lf(self, parent, title):
        return tk.LabelFrame(
            parent, text=f"  {title}  ",
            bg=C["bg2"], fg=C["cyan"],
            font=("Segoe UI", 9, "bold"),
            relief="flat", bd=0,
            highlightbackground=C["border"],
            highlightthickness=1,
        )

    def _build(self):
        PAD = dict(padx=10, pady=5)

        # ── Description ─────────────────────────────────────────────────────
        tk.Label(self,
            text="Recursively extracts ZIP, 7Z, and RAR archives into their own named subfolders. "
                 "Detects archives nested inside extracted content and can auto-extract them in the same pass. "
                 "Extracted archives can be kept, recycled, deleted, or moved to a separate location.",
            bg=C["bg2"], fg=C["muted"], font=("Segoe UI", 8), wraplength=860,
            justify="left", anchor="w", padx=12, pady=6
        ).pack(fill="x", padx=10, pady=(6, 0))

        # ── Source ──────────────────────────────────────────────────────────
        sf = self._lf(self, "Source Folder")
        sf.pack(fill="x", **PAD)
        self.src = FolderPicker(sf)
        self.src.pack(fill="x", padx=8, pady=6)

        # ── Destination ─────────────────────────────────────────────────────
        df = self._lf(self, "Destination")
        df.pack(fill="x", **PAD)

        mode_row = tk.Frame(df, bg=C["bg2"])
        mode_row.pack(fill="x", padx=8, pady=(6, 2))

        self.dst_mode = tk.StringVar(value="same")
        for text, val in [("Same as source", "same"), ("Mirror to custom destination", "custom")]:
            tk.Radiobutton(
                mode_row, text=text, variable=self.dst_mode, value=val,
                bg=C["bg2"], fg=C["text"], selectcolor=C["bg3"],
                activebackground=C["bg2"], font=("Segoe UI", 9),
                command=self._on_dst_mode
            ).pack(side="left", padx=(0, 14))

        self.dst = FolderPicker(df, label="Drop destination root here  —  or  Browse →",
                                disabled=True)
        self.dst.pack(fill="x", padx=8, pady=(2, 2))

        tk.Label(df,
                 text="  Mirror example:  D:\\source\\sub\\file.zip  →  E:\\dest\\source\\sub\\file\\",
                 bg=C["bg2"], fg=C["muted"], font=("Segoe UI", 8, "italic")
                 ).pack(anchor="w", padx=8, pady=(0, 4))

        # ── Options ─────────────────────────────────────────────────────────
        of = self._lf(self, "Options")
        of.pack(fill="x", **PAD)

        row1 = tk.Frame(of, bg=C["bg2"])
        row1.pack(fill="x", padx=8, pady=(6, 2))

        tk.Label(row1, text="Formats:", bg=C["bg2"], fg=C["text"],
                 font=("Segoe UI", 9)).pack(side="left")

        self.fmt_zip = tk.BooleanVar(value=True)
        self.fmt_7z  = tk.BooleanVar(value=True)
        self.fmt_rar = tk.BooleanVar(value=True)
        for var, lbl in [(self.fmt_zip, ".zip"), (self.fmt_7z, ".7z"), (self.fmt_rar, ".rar")]:
            tk.Checkbutton(
                row1, text=lbl, variable=var,
                bg=C["bg2"], fg=C["text"], selectcolor=C["bg3"],
                activebackground=C["bg2"], font=("Segoe UI", 9)
            ).pack(side="left", padx=(8, 0))

        self.recurse_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            row1, text="Recursive", variable=self.recurse_var,
            bg=C["bg2"], fg=C["text"], selectcolor=C["bg3"],
            activebackground=C["bg2"], font=("Segoe UI", 9)
        ).pack(side="left", padx=(24, 0))

        self.nested_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            row1, text="Auto-extract nested archives",
            variable=self.nested_var,
            bg=C["bg2"], fg=C["accent"], selectcolor=C["bg3"],
            activebackground=C["bg2"], font=("Segoe UI", 9, "bold")
        ).pack(side="left", padx=(24, 0))

        row2a = tk.Frame(of, bg=C["bg2"])
        row2a.pack(fill="x", padx=8, pady=(2, 0))

        tk.Label(row2a, text="After extraction:", bg=C["bg2"], fg=C["text"],
                 font=("Segoe UI", 9)).pack(side="left")

        self.after_mode = tk.StringVar(value="keep")
        rb_row1 = [
            ("Keep archive",      "keep",      C["text"]),
            ("→ Recycle Bin",     "recycle",   C["cyan"]),
            ("→ Permanent delete","permanent", C["amber"]),
        ]
        for txt, val, fg in rb_row1:
            state = "normal"
            if val == "recycle" and not TRASH_AVAILABLE:
                state = "disabled"
                txt += " (send2trash missing)"
            tk.Radiobutton(
                row2a, text=txt, variable=self.after_mode, value=val,
                bg=C["bg2"], fg=fg, selectcolor=C["bg3"],
                activebackground=C["bg2"], font=("Segoe UI", 9),
                state=state, command=self._on_after_mode
            ).pack(side="left", padx=(10, 0))

        row2b = tk.Frame(of, bg=C["bg2"])
        row2b.pack(fill="x", padx=8, pady=(2, 0))

        tk.Label(row2b, text=" " * 17, bg=C["bg2"]).pack(side="left")   # indent align
        rb_row2 = [
            ("→ Move (mirror structure)", "move_mirror", C["green"]),
            ("→ Move (flat dump)",        "move_flat",   C["green"]),
        ]
        for txt, val, fg in rb_row2:
            tk.Radiobutton(
                row2b, text=txt, variable=self.after_mode, value=val,
                bg=C["bg2"], fg=fg, selectcolor=C["bg3"],
                activebackground=C["bg2"], font=("Segoe UI", 9),
                command=self._on_after_mode
            ).pack(side="left", padx=(10, 0))

        # Move destination — shown only when a move mode is active
        self.move_dst_frame = tk.Frame(of, bg=C["bg2"])
        tk.Label(self.move_dst_frame,
                 text="  Move destination:",
                 bg=C["bg2"], fg=C["text"], font=("Segoe UI", 9)
                 ).pack(anchor="w", padx=8, pady=(4, 0))
        self.move_dst = FolderPicker(
            self.move_dst_frame,
            label="Drop move-destination folder here  —  or  Browse →"
        )
        self.move_dst.pack(fill="x", padx=8, pady=(2, 6))
        # Not packed initially; _on_after_mode shows/hides it

        # ── Stats + Progress ─────────────────────────────────────────────────
        stat_row = tk.Frame(self, bg=C["bg"])
        stat_row.pack(fill="x", padx=10, pady=(6, 1))
        self.stat_var = tk.StringVar(value="Ready.")
        tk.Label(stat_row, textvariable=self.stat_var, bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 8)).pack(side="left")

        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=10, pady=(1, 4))

        # ── Log ─────────────────────────────────────────────────────────────
        self.log = LogPane(self)
        self.log.pack(fill="both", expand=True, padx=10, pady=(2, 4))

        # ── Buttons ─────────────────────────────────────────────────────────
        btn = tk.Frame(self, bg=C["bg"])
        btn.pack(fill="x", padx=10, pady=(0, 10))

        self.start_btn = tk.Button(
            btn, text="▶   Extract", bg=C["green"], fg="#000",
            relief="flat", font=("Segoe UI", 10, "bold"),
            command=self._start, cursor="hand2", padx=20
        )
        self.start_btn.pack(side="left")

        self.stop_btn = tk.Button(
            btn, text="■   Stop", bg=C["red"], fg="white",
            relief="flat", font=("Segoe UI", 10, "bold"),
            command=lambda: setattr(self, "_stop", True),
            cursor="hand2", padx=20, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=(8, 0))

        tk.Button(btn, text="Save Log", bg=C["bg4"], fg=C["muted"],
                  relief="flat", font=("Segoe UI", 9),
                  command=self.log.save, cursor="hand2", padx=10
                  ).pack(side="right")
        tk.Button(btn, text="Clear Log", bg=C["bg4"], fg=C["muted"],
                  relief="flat", font=("Segoe UI", 9),
                  command=self.log.clear, cursor="hand2", padx=10
                  ).pack(side="right", padx=(0, 6))

    def _on_dst_mode(self):
        self.dst.enable(self.dst_mode.get() == "custom")

    def _on_after_mode(self):
        is_move = self.after_mode.get() in ("move_mirror", "move_flat")
        if is_move:
            self.move_dst_frame.pack(fill="x", padx=8, pady=(0, 6))
        else:
            self.move_dst_frame.pack_forget()

    def _stat(self, msg):
        self.after(0, lambda: self.stat_var.set(msg))

    def _prog(self, val):
        self.after(0, lambda: self.progress.config(value=val))

    def _start(self):
        src_path = self.src.get()
        if not src_path or not Path(src_path).is_dir():
            self.log.write("fail", "ERROR: Source folder not set or does not exist.\n")
            return

        custom_dst = None
        if self.dst_mode.get() == "custom":
            custom_dst = self.dst.get()
            if not custom_dst:
                self.log.write("fail", "ERROR: Custom destination not set.\n")
                return

        after = self.after_mode.get()
        move_root = None
        if after in ("move_mirror", "move_flat"):
            move_root = self.move_dst.get()
            if not move_root:
                self.log.write("fail", "ERROR: Move destination not set.\n")
                return
            move_root = Path(move_root)

        exts = set()
        if self.fmt_zip.get(): exts.add(".zip")
        if self.fmt_7z.get():  exts.add(".7z")
        if self.fmt_rar.get(): exts.add(".rar")
        if not exts:
            self.log.write("fail", "ERROR: No formats selected.\n")
            return

        self._stop = False
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        threading.Thread(
            target=self._run,
            args=(Path(src_path),
                  Path(custom_dst) if custom_dst else None,
                  exts,
                  after,
                  move_root,
                  self.recurse_var.get(),
                  self.nested_var.get()),
            daemon=True
        ).start()

    def _run(self, src_root: Path, dst_root, exts: set,
             after_mode: str, move_root, recurse: bool, auto_nested: bool):
        t0 = time.time()

        # Collect initial archives
        initial = []
        for ext in exts:
            initial.extend(src_root.rglob(f"*{ext}") if recurse else src_root.glob(f"*{ext}"))
        initial = sorted(set(initial))

        if not initial:
            self.log.write("info", f"No archives found under: {src_root}\n")
            self._finish()
            return

        queue      = deque(initial)
        queued_set = set(initial)   # track to avoid re-queuing the same file
        total_seen = len(initial)
        processed  = 0

        self.log.write("info", f"Found {total_seen} archive(s) under: {src_root}\n")
        if dst_root:
            self.log.write("info", f"Extract destination: {dst_root}\n")
        if move_root:
            mode_label = "mirror" if after_mode == "move_mirror" else "flat"
            self.log.write("info", f"Move destination ({mode_label}): {move_root}\n")
        if auto_nested:
            self.log.write("info", "Auto-extract nested: ON\n")
        self.log.write("info", "─" * 64 + "\n")

        ok = fail = bad = nested_total = 0

        while queue and not self._stop:
            arc = queue.popleft()
            processed += 1

            elapsed = time.time() - t0
            rate    = processed / elapsed if elapsed > 0 else 0
            remain  = len(queue)
            eta     = remain / rate if rate > 0 else 0
            self._stat(
                f"{processed}/{total_seen}  (+{remain} queued)  │  "
                f"✓ {ok}  ✗ {fail}  │  "
                f"{timedelta(seconds=int(elapsed))} elapsed  ETA {timedelta(seconds=int(eta))}"
            )
            self._prog(min(99, 100 * processed / total_seen))

            mode, _ = classify(arc)

            if mode == "bad":
                self.log.write("mute", f"[BAD]   {arc}\n")
                bad += 1
                continue

            # ── Resolve extraction output location ─────────────────────────
            # Nested archives (not under src_root) extract beside themselves
            try:
                rel_parent = arc.relative_to(src_root).parent
                under_src  = True
            except ValueError:
                rel_parent = Path(".")
                under_src  = False

            if dst_root is not None and under_src:
                if mode == "single":
                    out_dir = dst_root / src_root.name / rel_parent
                else:
                    out_dir = dst_root / src_root.name / rel_parent / sanitize(arc.stem)
            else:
                if mode == "single":
                    out_dir = arc.parent
                else:
                    out_dir = arc.parent / sanitize(arc.stem)

            # ── Extract ────────────────────────────────────────────────────
            if mode == "single":
                ok_ex, err = extract_single(arc, out_dir)
            else:
                ok_ex, err = extract_to_folder(arc, out_dir)

            if not ok_ex:
                self.log.write("fail", f"[FAIL]  {arc.name}\n        {err}\n")
                fail += 1
                continue

            # ── Scan output for nested archives ────────────────────────────
            nested_found = scan_for_archives(out_dir, exts)
            if nested_found:
                nested_total += len(nested_found)
                sep = "▼" * 60
                self.log.write("nested",
                    f"{sep}\n"
                    f"  ⚠  NESTED ARCHIVES DETECTED in: {out_dir.name}\n"
                    f"  ↳  {len(nested_found)} archive(s) found after extracting {arc.name}\n"
                )
                for nf in nested_found:
                    if auto_nested and nf not in queued_set:
                        queued_set.add(nf)
                        queue.append(nf)
                        total_seen += 1
                        self.log.write("nested", f"       [QUEUED]  {nf.name}\n")
                    else:
                        action = "(already queued)" if nf in queued_set else "(not auto-extracting)"
                        self.log.write("nested", f"       [FOUND]   {nf.name}  {action}\n")
                self.log.write("nested", f"{sep}\n")

            # ── Post-extract action ────────────────────────────────────────
            if after_mode == "keep":
                suffix, tag = "", "ok"

            elif after_mode == "recycle":
                ok_d, err_d = delete_archive(arc, "recycle")
                suffix = "  [recycled]" if ok_d else f"  [recycle WARN: {err_d}]"
                tag = "ok" if ok_d else "warn"

            elif after_mode == "permanent":
                ok_d, err_d = delete_archive(arc, "permanent")
                suffix = "  [deleted]" if ok_d else f"  [delete WARN: {err_d}]"
                tag = "ok" if ok_d else "warn"

            elif after_mode == "move_mirror":
                ok_d, dest_or_err = move_mirrored(arc, src_root, move_root)
                suffix = f"  [→ {dest_or_err}]" if ok_d else f"  [move WARN: {dest_or_err}]"
                tag = "ok" if ok_d else "warn"

            elif after_mode == "move_flat":
                ok_d, dest_or_err = move_flat(arc, move_root)
                suffix = f"  [→ {dest_or_err}]" if ok_d else f"  [move WARN: {dest_or_err}]"
                tag = "ok" if ok_d else "warn"

            else:
                suffix, tag = "", "ok"

            self.log.write(tag, f"[OK]    {arc.name}  →  {out_dir}{suffix}\n")
            ok += 1

        if self._stop:
            self.log.write("warn", f"[STOPPED — {len(queue)} remaining in queue]\n")

        elapsed = time.time() - t0
        self.log.write("info", "─" * 64 + "\n")
        nested_note = f"  │  Nested alerts: {nested_total}" if nested_total else ""
        self.log.write("info",
                       f"Done.  OK: {ok}  Fail: {fail}  Bad: {bad}{nested_note}  │  "
                       f"{timedelta(seconds=int(elapsed))}\n")
        self._stat(f"Done — OK: {ok}, Fail: {fail}, Bad: {bad}"
                   + (f", Nested alerts: {nested_total}" if nested_total else ""))
        self._prog(100)
        self._finish()

    def _finish(self):
        self.after(0, lambda: self.start_btn.config(state="normal"))
        self.after(0, lambda: self.stop_btn.config(state="disabled"))


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — ZIP STORE PACKER
# ══════════════════════════════════════════════════════════════════════════════

class ZipStorePackerTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=C["bg"])
        self._stop  = False
        self._exts  = []
        self._build()

    def _lf(self, parent, title):
        return tk.LabelFrame(
            parent, text=f"  {title}  ",
            bg=C["bg2"], fg=C["cyan"],
            font=("Segoe UI", 9, "bold"),
            relief="flat", bd=0,
            highlightbackground=C["border"],
            highlightthickness=1,
        )

    def _build(self):
        PAD = dict(padx=10, pady=5)

        # ── Description ─────────────────────────────────────────────────────
        tk.Label(self,
            text="Wraps files in uncompressed ZIP containers (ZIP_STORED — zero compression) "
                 "for use as a neutral byte-preserving wrapper before downstream recompression. "
                 "Each source file is verified inside its zip before the original is deleted. "
                 "Target extensions are configurable; existing zips are skipped by default.",
            bg=C["bg2"], fg=C["muted"], font=("Segoe UI", 8), wraplength=860,
            justify="left", anchor="w", padx=12, pady=6
        ).pack(fill="x", padx=10, pady=(6, 0))

        # ── Source ──────────────────────────────────────────────────────────
        sf = self._lf(self, "Target Folder")
        sf.pack(fill="x", **PAD)
        self.src = FolderPicker(sf)
        self.src.pack(fill="x", padx=8, pady=6)

        # ── Extensions ──────────────────────────────────────────────────────
        ef = self._lf(self, "Target Extensions")
        ef.pack(fill="x", **PAD)

        add_row = tk.Frame(ef, bg=C["bg2"])
        add_row.pack(fill="x", padx=8, pady=(6, 4))

        self.ext_entry = tk.Entry(
            add_row, bg=C["bg3"], fg=C["text"], insertbackground=C["text"],
            relief="flat", font=("Segoe UI", 9), width=22
        )
        self.ext_entry.pack(side="left")
        self.ext_entry.insert(0, "exe")
        self.ext_entry.bind("<Return>", lambda e: self._add_ext())

        tk.Button(
            add_row, text="Add", bg=C["accent"], fg="white",
            relief="flat", font=("Segoe UI", 9, "bold"),
            command=self._add_ext, cursor="hand2", padx=10
        ).pack(side="left", padx=(6, 0))

        tk.Label(add_row, text="space/comma separated — e.g.  exe dll bin rom",
                 bg=C["bg2"], fg=C["muted"], font=("Segoe UI", 8, "italic")
                 ).pack(side="left", padx=10)

        self.pill_frame = tk.Frame(ef, bg=C["bg2"])
        self.pill_frame.pack(fill="x", padx=8, pady=(0, 6))
        self._add_ext_internal(".exe")   # default

        # ── Options ─────────────────────────────────────────────────────────
        of = self._lf(self, "Options")
        of.pack(fill="x", **PAD)

        row = tk.Frame(of, bg=C["bg2"])
        row.pack(fill="x", padx=8, pady=6)

        self.recurse_var = tk.BooleanVar(value=True)
        self.verify_var  = tk.BooleanVar(value=True)
        self.skip_var    = tk.BooleanVar(value=True)

        for var, lbl in [
            (self.recurse_var, "Recursive"),
            (self.verify_var,  "Verify before delete"),
            (self.skip_var,    "Skip if .zip already exists"),
        ]:
            tk.Checkbutton(
                row, text=lbl, variable=var,
                bg=C["bg2"], fg=C["text"], selectcolor=C["bg3"],
                activebackground=C["bg2"], font=("Segoe UI", 9)
            ).pack(side="left", padx=(0, 18))

        # ── Stats + Progress ─────────────────────────────────────────────────
        stat_row = tk.Frame(self, bg=C["bg"])
        stat_row.pack(fill="x", padx=10, pady=(6, 1))
        self.stat_var = tk.StringVar(value="Ready.")
        tk.Label(stat_row, textvariable=self.stat_var, bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 8)).pack(side="left")

        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=10, pady=(1, 4))

        # ── Log ─────────────────────────────────────────────────────────────
        self.log = LogPane(self)
        self.log.pack(fill="both", expand=True, padx=10, pady=(2, 4))

        # ── Buttons ─────────────────────────────────────────────────────────
        btn = tk.Frame(self, bg=C["bg"])
        btn.pack(fill="x", padx=10, pady=(0, 10))

        self.start_btn = tk.Button(
            btn, text="▶   Pack", bg=C["green"], fg="#000",
            relief="flat", font=("Segoe UI", 10, "bold"),
            command=self._start, cursor="hand2", padx=20
        )
        self.start_btn.pack(side="left")

        self.stop_btn = tk.Button(
            btn, text="■   Stop", bg=C["red"], fg="white",
            relief="flat", font=("Segoe UI", 10, "bold"),
            command=lambda: setattr(self, "_stop", True),
            cursor="hand2", padx=20, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=(8, 0))

        tk.Button(btn, text="Save Log", bg=C["bg4"], fg=C["muted"],
                  relief="flat", font=("Segoe UI", 9),
                  command=self.log.save, cursor="hand2", padx=10
                  ).pack(side="right")
        tk.Button(btn, text="Clear Log", bg=C["bg4"], fg=C["muted"],
                  relief="flat", font=("Segoe UI", 9),
                  command=self.log.clear, cursor="hand2", padx=10
                  ).pack(side="right", padx=(0, 6))

    # ── Extension pill management ──────────────────────────────────────────

    def _add_ext(self):
        raw = self.ext_entry.get()
        for tok in re.split(r"[\s,]+", raw):
            tok = tok.strip().lstrip(".")
            if tok:
                self._add_ext_internal("." + tok.lower())
        self.ext_entry.delete(0, "end")

    def _add_ext_internal(self, ext: str):
        if ext in self._exts:
            return
        self._exts.append(ext)

        pill = tk.Frame(self.pill_frame, bg=C["accent"], padx=6, pady=2)
        pill.pack(side="left", padx=(0, 5), pady=3)
        tk.Label(pill, text=ext, bg=C["accent"], fg="white",
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        x = tk.Label(pill, text=" ✕", bg=C["accent"], fg="white",
                     font=("Segoe UI", 9), cursor="hand2")
        x.pack(side="left")
        x.bind("<Button-1>", lambda e, p=pill, ex=ext: self._remove_ext(p, ex))

    def _remove_ext(self, pill: tk.Frame, ext: str):
        if ext in self._exts:
            self._exts.remove(ext)
        pill.destroy()

    # ── Run ───────────────────────────────────────────────────────────────

    def _stat(self, msg):
        self.after(0, lambda: self.stat_var.set(msg))

    def _prog(self, val):
        self.after(0, lambda: self.progress.config(value=val))

    def _start(self):
        src_path = self.src.get()
        if not src_path or not Path(src_path).is_dir():
            self.log.write("fail", "ERROR: Target folder not set or does not exist.\n")
            return
        if not self._exts:
            self.log.write("fail", "ERROR: No extensions configured.\n")
            return

        self._stop = False
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        threading.Thread(
            target=self._run,
            args=(Path(src_path), list(self._exts),
                  self.recurse_var.get(), self.verify_var.get(), self.skip_var.get()),
            daemon=True
        ).start()

    def _run(self, src: Path, exts: list, recurse: bool, verify: bool, skip_existing: bool):
        t0 = time.time()

        files = []
        for ext in exts:
            files.extend(src.rglob(f"*{ext}") if recurse else src.glob(f"*{ext}"))
        files = sorted(set(files))

        total = len(files)
        if total == 0:
            self.log.write("info", f"No matching files found under: {src}\n")
            self._finish()
            return

        self.log.write("info",
                       f"Found {total} file(s) under: {src}\n"
                       f"Extensions: {', '.join(exts)}\n")
        self.log.write("info", "─" * 64 + "\n")

        ok = fail = skipped = 0

        for i, fp in enumerate(files, 1):
            if self._stop:
                self.log.write("warn", f"[STOPPED at {i-1}/{total}]\n")
                break

            elapsed = time.time() - t0
            rate    = i / elapsed if elapsed > 0 else 0
            eta     = (total - i) / rate if rate > 0 else 0
            self._stat(
                f"{i}/{total}  │  ✓ {ok}  ✗ {fail}  ↷ {skipped}  │  "
                f"{timedelta(seconds=int(elapsed))} elapsed  ETA {timedelta(seconds=int(eta))}"
            )
            self._prog(100 * i / total)

            zip_path = fp.with_suffix(".zip")

            if skip_existing and zip_path.exists():
                self.log.write("skip", f"[SKIP]  {fp.name}  (zip already exists)\n")
                skipped += 1
                continue

            # ── Create ZIP_STORED ──────────────────────────────────────────
            try:
                with zipfile.ZipFile(zip_path, "w",
                                     compression=zipfile.ZIP_STORED,
                                     allowZip64=True) as zf:
                    zf.write(fp, fp.name)
            except Exception as e:
                self.log.write("fail", f"[FAIL]  {fp.name}: create: {e}\n")
                zip_path.unlink(missing_ok=True)
                fail += 1
                continue

            # ── Verify ────────────────────────────────────────────────────
            if verify:
                try:
                    with zipfile.ZipFile(zip_path, "r") as zf:
                        bad = zf.testzip()
                        if bad:
                            raise ValueError(f"corrupt entry: {bad}")
                        info = zf.getinfo(fp.name)
                        if info.file_size != fp.stat().st_size:
                            raise ValueError(
                                f"size mismatch ({info.file_size} vs {fp.stat().st_size})")
                except Exception as e:
                    self.log.write("fail", f"[FAIL]  {fp.name}: verify: {e}\n")
                    zip_path.unlink(missing_ok=True)
                    fail += 1
                    continue

            # ── Delete original ────────────────────────────────────────────
            try:
                fp.unlink()
                sz = zip_path.stat().st_size
                self.log.write("ok", f"[OK]    {fp.name}  ({sz:,} B)\n")
                ok += 1
            except Exception as e:
                self.log.write("warn", f"[WARN]  {fp.name}: packed OK, delete failed: {e}\n")
                ok += 1

        elapsed = time.time() - t0
        self.log.write("info", "─" * 64 + "\n")
        self.log.write("info",
                       f"Done.  OK: {ok}  Fail: {fail}  Skip: {skipped}  │  "
                       f"{timedelta(seconds=int(elapsed))}\n")
        self._stat(f"Done — OK: {ok}, Fail: {fail}, Skip: {skipped}")
        self._prog(100)
        self._finish()

    def _finish(self):
        self.after(0, lambda: self.start_btn.config(state="normal"))
        self.after(0, lambda: self.stop_btn.config(state="disabled"))


# ══════════════════════════════════════════════════════════════════════════════
#  APPLICATION SHELL
# ══════════════════════════════════════════════════════════════════════════════

class App:
    def __init__(self):
        Root = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk
        self.root = Root()
        self.root.title("Eggman's Archive Utilities  v2.0")
        self.root.geometry("900x900")
        self.root.minsize(900, 600)
        self.root.configure(bg=C["bg"])

        self._apply_style()
        self._build_header()
        self._build_notebook()

    def _apply_style(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TNotebook",        background=C["bg"],  borderwidth=0)
        s.configure("TNotebook.Tab",    background=C["bg3"], foreground=C["muted"],
                    font=("Segoe UI", 10, "bold"), padding=[18, 7])
        s.map("TNotebook.Tab",
              background=[("selected", C["bg2"])],
              foreground=[("selected", C["cyan"])])
        s.configure("Horizontal.TProgressbar",
                    troughcolor=C["bg3"], background=C["accent"],
                    borderwidth=0, thickness=5)

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C["bg2"], pady=9)
        hdr.pack(fill="x")

        tk.Label(hdr, text="⚙  Eggman's Archive Utilities",
                 bg=C["bg2"], fg=C["cyan"],
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=16)
        tk.Label(hdr, text="v2.0",
                 bg=C["bg2"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left", pady=2)

        warnings = []
        if not DND_AVAILABLE: warnings.append("pip install tkinterdnd2  (drag-and-drop)")
        if not TRASH_AVAILABLE: warnings.append("pip install send2trash  (recycle bin)")
        if warnings:
            tk.Label(hdr, text="⚠  " + "   ·   ".join(warnings),
                     bg=C["bg2"], fg=C["amber"],
                     font=("Segoe UI", 8)).pack(side="right", padx=16)

    def _build_notebook(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)
        nb.add(ExtractorTab(nb),      text="  📦  Extractor  ")
        nb.add(ZipStorePackerTab(nb), text="  🗜  ZIP Store Packer  ")

    def run(self):
        self.root.mainloop()


# ══════════════════════════════════════════════════════════════════════════════

def main():
    try:
        ensure_7z()
    except FileNotFoundError as e:
        import tkinter.messagebox as mb
        root = tk.Tk(); root.withdraw()
        mb.showerror("7z not found", str(e))
        sys.exit(1)
    App().run()


if __name__ == "__main__":
    main()
