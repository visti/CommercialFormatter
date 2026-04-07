"""Customtkinter GUI for Commercial Formatter — deep-space aesthetic."""
from __future__ import annotations

import io
import math
import os
import queue
import re
import sys
import threading
import tomllib
from pathlib import Path
from typing import Any, Callable

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkfont

try:
    from PIL import Image, ImageDraw, ImageFilter
    from PIL.ImageTk import PhotoImage as PILPhotoImage
    _PIL = True
except ImportError:
    _PIL = False

from .decisions import DecisionConfig, Option
from .interaction import InteractionHandler
from .output import ProcessingStats, print_summary_box
from .processor import generate_rejection_filename, get_files, process_files, read_files
from .stations import STATIONS, get_station, list_stations

# ── Appearance ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Design tokens ─────────────────────────────────────────────────────────────
BG_BASE    = "#050506"
BG_RAISED  = "#0a0a0d"
BG_SURFACE = "#0f0f14"
SIDEBAR_BG = "#0c0c10"

ACCENT     = "#5E6AD2"
ACCENT_HI  = "#6872D9"
ACCENT_DIM = "#3a4296"

FG        = "#EDEDEF"
FG_MUTED  = "#8A8F98"
FG_SUBTLE = "#3e424a"

BORDER = "#1c1c24"

C_SUCCESS = "#3aba3a"
C_ERROR   = "#c03030"
C_WARN    = "#d4a020"

_ANSI_RE = re.compile(r"\033\[[0-9;]*m")

# Floating-panel layout
MARGIN = 14   # gap between window edge / panels
SIDE_W = 282  # sidebar width


# ── Font detection (cached) ───────────────────────────────────────────────────
_PROP: str | None = None
_MONO: str | None = None


def _prop_font() -> str:
    global _PROP
    if _PROP is None:
        try:
            fams = set(tkfont.families())
            for f in ("Inter", "Segoe UI", "SF Pro Display", "Helvetica Neue"):
                if f in fams:
                    _PROP = f; return f
        except Exception:
            pass
        _PROP = "Helvetica"
    return _PROP


def _mono_font() -> str:
    global _MONO
    if _MONO is None:
        try:
            fams = set(tkfont.families())
            for f in ("Consolas", "JetBrains Mono", "SF Mono", "Menlo", "Courier New"):
                if f in fams:
                    _MONO = f; return f
        except Exception:
            pass
        _MONO = "Courier"
    return _MONO


# ── Color interpolation ───────────────────────────────────────────────────────
def _lerp(c1: str, c2: str, t: float) -> str:
    r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
    r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
    return "#{:02x}{:02x}{:02x}".format(
        max(0, min(255, round(r1 + (r2 - r1) * t))),
        max(0, min(255, round(g1 + (g2 - g1) * t))),
        max(0, min(255, round(b1 + (b2 - b1) * t))),
    )


# ── Config helpers ────────────────────────────────────────────────────────────
def _folder_mapping() -> dict[str, str]:
    p = Path(__file__).parent.parent / "config" / "folders.toml"
    if not p.exists():
        return {}
    try:
        with open(p, "rb") as f:
            return tomllib.load(f).get("folders", {})
    except Exception:
        return {}


# ── Animated PIL background ───────────────────────────────────────────────────
if _PIL:
    class _BgRenderer:
        """Gradient base + three lazily floating accent blobs."""
        _S = 6  # blob overlay scale divisor

        def __init__(self, w: int, h: int) -> None:
            self.w, self.h, self.phase = w, h, 0.0
            self._base = self._gradient(w, h).convert("RGBA")

        # ── private ──────────────────────────────────────────────────────
        @staticmethod
        def _gradient(w: int, h: int) -> Image.Image:
            """1-pixel-wide gradient, scaled to full width."""
            col = Image.new("RGB", (1, max(h, 1)))
            d   = ImageDraw.Draw(col)
            for y in range(h):
                t = y / max(h - 1, 1)
                r = max(0, round(10 - 8 * t))
                g = max(0, round(10 - 8 * t))
                b = max(0, round(15 - 12 * t))
                d.point((0, y), fill=(r, g, b))
            return col.resize((w, h), Image.NEAREST)

        def next_frame(self) -> Image.Image:
            S  = self._S
            sw = max(self.w // S, 60)
            sh = max(self.h // S, 40)
            p  = self.phase

            ov = Image.new("RGBA", (sw, sh), (0, 0, 0, 0))
            d  = ImageDraw.Draw(ov)

            # Blob 1 — indigo, top-right
            x1 = sw * .68 + math.sin(p)           * sw * .05
            y1 = sh * .08 + math.cos(p * .75)     * sh * .06
            d.ellipse([x1-sw*.33, y1-sh*.50, x1+sw*.33, y1+sh*.50],
                      fill=(94, 106, 210, 55))

            # Blob 2 — violet, left-centre
            x2 = sw * .07 + math.cos(p * .60 + 1.4) * sw * .04
            y2 = sh * .62 + math.sin(p * .45 + .90) * sh * .06
            d.ellipse([x2-sw*.20, y2-sh*.32, x2+sw*.20, y2+sh*.32],
                      fill=(60, 40, 160, 35))

            # Blob 3 — teal, bottom-centre
            x3 = sw * .45 + math.sin(p * .35 + 2.1) * sw * .06
            y3 = sh * .90
            d.ellipse([x3-sw*.22, y3-sh*.18, x3+sw*.22, y3+sh*.18],
                      fill=(40, 80, 150, 22))

            ov = ov.filter(ImageFilter.GaussianBlur(radius=max(sh // 4, 6)))
            ov = ov.resize((self.w, self.h), Image.BILINEAR)

            self.phase += 0.012
            return Image.alpha_composite(self._base, ov).convert("RGB")

        def resize(self, w: int, h: int) -> None:
            if (w, h) != (self.w, self.h):
                self.w, self.h = w, h
                self._base = self._gradient(w, h).convert("RGBA")


# ── Log writer ────────────────────────────────────────────────────────────────
class LogWriter(io.RawIOBase):
    """Thread-safe stdout that strips ANSI and feeds a queue."""

    def __init__(self, q: queue.Queue[str]) -> None:
        self._q = q

    def write(self, text: str) -> int:  # type: ignore[override]
        clean = _ANSI_RE.sub("", text)
        if clean:
            self._q.put(clean)
        return len(text)

    def writable(self) -> bool: return True
    def flush(self)    -> None: pass
    def isatty(self)   -> bool: return False


# ── GUI interaction handler ───────────────────────────────────────────────────
class GUIInteractionHandler(InteractionHandler):
    """Blocks the processing thread and shows a CTkToplevel dialog."""

    def __init__(self, app: App) -> None:
        self._app    = app
        self._result: Any = None
        self._event  = threading.Event()

    def prompt_decision(
        self,
        config: DecisionConfig,
        key: Any = None,
        count: int = 0,
        extra_input_fn: Callable[[str], Any] | None = None,
    ) -> tuple[str, Any]:
        self._result = None
        self._event.clear()
        self._app.after(0, lambda: DecisionDialog(
            self._app, config, key, count, self._resolve))
        self._event.wait()
        return self._result  # type: ignore[return-value]

    def prompt_string(self, message: str, default: str = "") -> str:
        self._result = None
        self._event.clear()
        self._app.after(0, lambda: StringInputDialog(
            self._app, message, default, self._resolve_str))
        self._event.wait()
        return self._result or default

    def _resolve(self, action: str, extra: Any = None) -> None:
        self._result = (action, extra)
        self._event.set()

    def _resolve_str(self, value: str) -> None:
        self._result = value
        self._event.set()


# ── Decision dialog ───────────────────────────────────────────────────────────
class DecisionDialog(ctk.CTkToplevel):
    _COLORS: dict[str, tuple[str, str]] = {
        "fix":    ("#1e4d1e", "#2a6a2a"),
        "accept": ("#1e4d1e", "#2a6a2a"),
        "keep":   ("#1e4d1e", "#2a6a2a"),
        "reject": ("#4d1e1e", "#6a2a2a"),
        "skip":   (BG_SURFACE, BG_RAISED),
        "edit":   (ACCENT_DIM, ACCENT),
    }

    def __init__(
        self,
        parent: ctk.CTk,
        config: DecisionConfig,
        key: Any,
        count: int,
        callback: Callable[[str, Any], None],
    ) -> None:
        super().__init__(parent)
        self.configure(fg_color=BG_SURFACE)
        self._config   = config
        self._key      = key
        self._count    = count
        self._callback = callback

        self.title("")
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build()
        self._center(parent)
        self.focus_force()

    # ── build ─────────────────────────────────────────────────────────────
    def _build(self) -> None:
        ff, fm = _prop_font(), _mono_font()

        # Header strip
        hdr = ctk.CTkFrame(self, fg_color=BG_RAISED, height=40, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="Decision Required",
            text_color=FG_MUTED, anchor="w",
            font=ctk.CTkFont(family=ff, size=11),
        ).pack(side="left", padx=16)

        # Issue detail
        detail = self._fmt_key()
        if detail:
            ctk.CTkLabel(
                self, text=detail,
                justify="left", anchor="w", wraplength=360,
                text_color=FG,
                font=ctk.CTkFont(family=fm, size=12),
            ).pack(fill="x", padx=18, pady=(14, 4))

        if self._count > 0:
            ctk.CTkLabel(
                self, text=f"{self._count}\u00d7",
                text_color=FG_SUBTLE, anchor="w",
                font=ctk.CTkFont(family=ff, size=11),
            ).pack(fill="x", padx=18, pady=(0, 6))

        ctk.CTkFrame(self, height=1, fg_color=BORDER,
                     corner_radius=0).pack(fill="x", pady=2)

        # Option buttons
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=12)
        for opt in self._config.options:
            fg, hov = self._COLORS.get(opt.action, (ACCENT_DIM, ACCENT))
            bw = 1 if opt.action == "skip" else 0
            ctk.CTkButton(
                row, text=opt.label, width=108, height=34,
                fg_color=fg, hover_color=hov,
                border_width=bw, border_color=BORDER,
                text_color=FG,
                font=ctk.CTkFont(family=ff, size=12),
                command=lambda o=opt: self._pick(o),
            ).pack(side="left", padx=3)

        self._extra = ctk.CTkFrame(self, fg_color="transparent")
        self._extra.pack(fill="x", padx=18, pady=(0, 12))

    # ── key display ───────────────────────────────────────────────────────
    def _fmt_key(self) -> str:
        name, k = self._config.name, self._key
        if k is None:
            return ""
        if name == "artist_title" and isinstance(k, tuple) and len(k) >= 2:
            title, artist = k[0], k[1]
            parts = title.split(" - ", 1)
            if len(parts) == 2:
                return (
                    f'Current:\n  Title:  "{title}"\n  Artist: "{artist}"\n\n'
                    f'Fixed:\n  Title:  "{parts[1]}"\n'
                    f'  Artist: "{artist}-{parts[0]}"'
                )
            return f'Title:  "{title}"\nArtist: "{artist}"'
        if name == "long_time" and isinstance(k, tuple) and len(k) >= 3:
            return f'Title:  "{k[0]}"\nArtist: "{k[1]}"\nTime:   {k[2]}'
        if name == "duplicate" and isinstance(k, tuple) and len(k) >= 3:
            return f'Title:  "{k[0]}"\nArtist: "{k[1]}"\nDate:   {k[2]}'
        if name == "multiple_years" and isinstance(k, dict):
            total = sum(k.values())
            lines = ["Dates from multiple years:"]
            for yr, cnt in sorted(k.items()):
                pct = f"  ({cnt / total * 100:.1f}%)" if total else ""
                lines.append(f"  {yr}:  {cnt:,}{pct}")
            return "\n".join(lines)
        return ""

    # ── interaction ───────────────────────────────────────────────────────
    def _pick(self, opt: Option) -> None:
        if opt.action == "edit":
            self._show_edit()
        else:
            self._callback(opt.action, None)
            self.destroy()

    def _show_edit(self) -> None:
        for w in self._extra.winfo_children():
            w.destroy()
        ff = _prop_font()
        ctk.CTkLabel(
            self._extra, text="Corrected time (MM:SS):",
            text_color=FG_MUTED,
            font=ctk.CTkFont(family=ff, size=11),
        ).pack(side="left", padx=(0, 6))
        e = ctk.CTkEntry(
            self._extra, width=80, placeholder_text="04:30",
            fg_color=BG_SURFACE, border_color=BORDER, text_color=FG,
        )
        e.pack(side="left", padx=4)
        e.focus()

        def confirm() -> None:
            val   = e.get().strip()
            parts = val.split(":")
            if len(parts) != 2:
                e.configure(border_color=C_ERROR); return
            try:
                int(parts[0]); int(parts[1])
            except ValueError:
                e.configure(border_color=C_ERROR); return
            self._callback("edit", val)
            self.destroy()

        e.bind("<Return>", lambda _: confirm())
        ctk.CTkButton(
            self._extra, text="OK", width=52, height=30,
            fg_color=ACCENT, hover_color=ACCENT_HI, text_color=FG,
            font=ctk.CTkFont(family=ff, size=12),
            command=confirm,
        ).pack(side="left", padx=4)

    def _on_close(self) -> None:
        default = next(
            (o.action for o in self._config.options if o.is_default),
            self._config.options[0].action,
        )
        self._callback(default, None)
        self.destroy()

    def _center(self, parent: ctk.CTk) -> None:
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h   = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")


# ── String input dialog ───────────────────────────────────────────────────────
class StringInputDialog(ctk.CTkToplevel):
    def __init__(
        self,
        parent: ctk.CTk,
        message: str,
        default: str,
        callback: Callable[[str], None],
    ) -> None:
        super().__init__(parent)
        self.configure(fg_color=BG_SURFACE)
        self._callback = callback
        self._default  = default
        self.title("")
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: self._submit(default))

        ff = _prop_font()
        hdr = ctk.CTkFrame(self, fg_color=BG_RAISED, height=40, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="Input Required",
            text_color=FG_MUTED, anchor="w",
            font=ctk.CTkFont(family=ff, size=11),
        ).pack(side="left", padx=16)

        ctk.CTkLabel(
            self, text=message,
            text_color=FG, anchor="w",
            font=ctk.CTkFont(family=ff, size=12),
        ).pack(fill="x", padx=18, pady=(14, 4))

        self._e = ctk.CTkEntry(
            self, width=320,
            fg_color=BG_SURFACE, border_color=BORDER, text_color=FG,
        )
        self._e.pack(padx=18, pady=4)
        if default:
            self._e.insert(0, default)
        self._e.focus()
        self._e.select_range(0, "end")

        ctk.CTkButton(
            self, text="Confirm", width=120, height=34,
            fg_color=ACCENT, hover_color=ACCENT_HI, text_color=FG,
            font=ctk.CTkFont(family=ff, size=12),
            command=lambda: self._submit(self._e.get()),
        ).pack(pady=(8, 14))
        self._e.bind("<Return>", lambda _: self._submit(self._e.get()))
        self._center(parent)

    def _submit(self, v: str) -> None:
        self._callback(v.strip() or self._default)
        self.destroy()

    def _center(self, parent: ctk.CTk) -> None:
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h   = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")


# ── Main window ───────────────────────────────────────────────────────────────
class App(ctk.CTk):
    _LOG_MS  = 80    # log-poll interval
    _BG_MS   = 120   # background frame interval
    _ANIM_MS = 50    # UI animation tick

    def __init__(self) -> None:
        super().__init__()
        self.title("Commercial Formatter")
        self.geometry("1100x700")
        self.minsize(820, 540)
        self.configure(fg_color=BG_BASE)

        self._log_q:   queue.Queue[str] = queue.Queue()
        self._processing  = False
        self._orig_stdout: Any = None
        self._anim_tick   = 0
        self._auto_filling = False
        self._output_modified = False
        self._last_wh = (0, 0)
        self._last_reject_path: Path | None = None

        # ── Animated gradient canvas (behind all panels) ─────────────────
        if _PIL:
            self._bg_canvas = tk.Canvas(self, highlightthickness=0, bg=BG_BASE)
            self._bg_canvas.place(x=0, y=0, relwidth=1, relheight=1)
            self._bg_item  = self._bg_canvas.create_image(0, 0, anchor="nw")
            self._bg_photo: Any = None
            self._bg_rend:  Any = None
            self.after(200, self._tick_bg)

        self.bind("<Configure>", self._on_resize)
        self._build_ui()
        self._populate_stations()
        self.after(self._LOG_MS, self._tick_log)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        M, SW = MARGIN, SIDE_W

        # Sidebar — fixed-width floating card
        self._sidebar = ctk.CTkFrame(
            self, fg_color=SIDEBAR_BG, width=SW,
            corner_radius=10, border_width=1, border_color=BORDER,
        )
        self._sidebar.pack(side="left", fill="y", padx=(M, 0), pady=M)
        self._sidebar.pack_propagate(False)
        self._build_sidebar(self._sidebar)

        # Log panel — fills all remaining space
        self._logpanel = ctk.CTkFrame(
            self, fg_color=BG_RAISED,
            corner_radius=10, border_width=1, border_color=BORDER,
        )
        self._logpanel.pack(side="left", fill="both", expand=True, padx=M, pady=M)
        self._build_logpanel(self._logpanel)

    # ── Sidebar ───────────────────────────────────────────────────────────────

    def _build_sidebar(self, p: ctk.CTkFrame) -> None:
        ff = _prop_font()
        p.grid_columnconfigure(0, weight=1)
        p.grid_rowconfigure(15, weight=1)  # spacer before run area

        # ── Title ─────────────────────────────────────────────────────────
        ctk.CTkLabel(
            p, text="Commercial Formatter",
            text_color=FG, anchor="w",
            font=ctk.CTkFont(family=ff, size=15, weight="bold"),
        ).grid(row=0, column=0, padx=18, pady=(18, 1), sticky="w")
        ctk.CTkLabel(
            p, text="Broadcast Metadata Processor",
            text_color=FG_SUBTLE, anchor="w",
            font=ctk.CTkFont(family=ff, size=10),
        ).grid(row=1, column=0, padx=18, pady=(0, 12), sticky="w")

        # ── INPUT ─────────────────────────────────────────────────────────
        self._shdr(p, "INPUT", row=2)

        self._flbl(p, "Working Directory", row=3)
        dr = ctk.CTkFrame(p, fg_color="transparent")
        dr.grid(row=4, column=0, padx=18, pady=(2, 10), sticky="ew")
        dr.grid_columnconfigure(0, weight=1)
        self._dir_var = ctk.StringVar(value=str(Path.cwd()))
        self._dir_var.trace_add("write", self._on_dir_typed)
        ctk.CTkEntry(
            dr, textvariable=self._dir_var, height=32,
            fg_color=BG_SURFACE, border_color=BORDER, text_color=FG,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            dr, text="\u2026", width=34, height=32,
            fg_color=BG_SURFACE, hover_color=BG_RAISED,
            border_width=1, border_color=BORDER, text_color=FG_MUTED,
            command=self._on_browse,
        ).grid(row=0, column=1)

        self._flbl(p, "Station", row=5)
        self._station_var = ctk.StringVar()
        self._station_combo = ctk.CTkComboBox(
            p, variable=self._station_var, height=32, state="readonly",
            fg_color=BG_SURFACE, border_color=BORDER,
            button_color=BG_SURFACE, button_hover_color=BG_RAISED,
            dropdown_fg_color=BG_SURFACE, dropdown_hover_color=BG_RAISED,
            text_color=FG,
            command=self._on_station_change,
        )
        self._station_combo.grid(row=6, column=0, padx=18, pady=(2, 8), sticky="ew")
        self._station_keys: list[str] = []

        # ── OUTPUT ────────────────────────────────────────────────────────
        self._shdr(p, "OUTPUT", row=7)

        self._flbl(p, "Filename", row=8)
        self._output_var = ctk.StringVar()
        self._output_var.trace_add("write", self._on_output_typed)
        ctk.CTkEntry(
            p, textvariable=self._output_var, height=32,
            fg_color=BG_SURFACE, border_color=BORDER, text_color=FG,
            placeholder_text="e.g. 2025_q4_abc.csv",
        ).grid(row=9, column=0, padx=18, pady=(2, 8), sticky="ew")

        # ── OPTIONS ───────────────────────────────────────────────────────
        self._shdr(p, "OPTIONS", row=10)

        self._no_sw_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            p, text="Disable stopword filtering",
            variable=self._no_sw_var,
            text_color=FG_MUTED, hover_color=BG_RAISED,
            font=ctk.CTkFont(family=ff, size=12),
            checkmark_color=FG, fg_color=ACCENT, border_color=BORDER,
        ).grid(row=11, column=0, padx=18, pady=(4, 3), sticky="w")

        self._no_rej_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            p, text="Skip rejection log file",
            variable=self._no_rej_var,
            text_color=FG_MUTED, hover_color=BG_RAISED,
            font=ctk.CTkFont(family=ff, size=12),
            checkmark_color=FG, fg_color=ACCENT, border_color=BORDER,
        ).grid(row=12, column=0, padx=18, pady=(0, 3), sticky="w")

        self._flbl(p, "Additional Filter", row=13)
        self._add_var = ctk.StringVar()
        ctk.CTkEntry(
            p, textvariable=self._add_var, height=32,
            fg_color=BG_SURFACE, border_color=BORDER, text_color=FG,
            placeholder_text="e.g. Spotify",
        ).grid(row=14, column=0, padx=18, pady=(2, 8), sticky="ew")

        # row 15 = spacer (weight=1)

        # ── Run zone ──────────────────────────────────────────────────────
        ctk.CTkFrame(
            p, height=1, fg_color=BORDER, corner_radius=0,
        ).grid(row=16, column=0, sticky="ew", padx=0, pady=0)

        self._run_btn = ctk.CTkButton(
            p, text="Run",
            height=48, corner_radius=0,
            fg_color=ACCENT, hover_color=ACCENT_HI,
            text_color=FG, border_width=0,
            font=ctk.CTkFont(family=ff, size=13, weight="bold"),
            command=self._on_run,
        )
        self._run_btn.grid(row=17, column=0, padx=0, pady=0, sticky="ew")

        self._status_var = ctk.StringVar(value="  \u25cf  Ready")
        self._status_lbl = ctk.CTkLabel(
            p, textvariable=self._status_var,
            text_color=FG_SUBTLE, anchor="w",
            font=ctk.CTkFont(family=ff, size=11),
        )
        self._status_lbl.grid(row=18, column=0, padx=18, pady=(6, 14), sticky="w")

    # ── Log panel ─────────────────────────────────────────────────────────────

    def _build_logpanel(self, p: ctk.CTkFrame) -> None:
        ff = _prop_font()
        p.grid_columnconfigure(0, weight=1)
        p.grid_rowconfigure(2, weight=1)

        hdr = ctk.CTkFrame(p, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
        hdr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            hdr, text="Log",
            text_color=FG, anchor="w",
            font=ctk.CTkFont(family=ff, size=12, weight="bold"),
        ).grid(row=0, column=0, sticky="w")
        self._open_rej_btn = ctk.CTkButton(
            hdr, text="Open reject file", width=110, height=24,
            fg_color="transparent", border_width=1, border_color=BORDER,
            text_color=FG_MUTED, hover_color=BG_SURFACE,
            font=ctk.CTkFont(family=ff, size=11),
            command=self._open_reject_file,
        )
        self._open_rej_btn.grid(row=0, column=1, sticky="e", padx=(0, 6))
        self._open_rej_btn.grid_remove()
        ctk.CTkButton(
            hdr, text="Clear", width=56, height=24,
            fg_color="transparent", border_width=1, border_color=BORDER,
            text_color=FG_MUTED, hover_color=BG_SURFACE,
            font=ctk.CTkFont(family=ff, size=11),
            command=self._clear_log,
        ).grid(row=0, column=2, sticky="e")

        ctk.CTkFrame(
            p, height=1, fg_color=BORDER, corner_radius=0,
        ).grid(row=1, column=0, sticky="ew", padx=0)

        self._log = ctk.CTkTextbox(
            p,
            fg_color="#060608",
            text_color=FG,
            font=ctk.CTkFont(family=_mono_font(), size=12),
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=FG_SUBTLE,
            wrap="none",
            corner_radius=0,
            border_width=0,
        )
        self._log.grid(row=2, column=0, sticky="nsew")
        self._log.configure(state="disabled")

    # ── Section / field-label helpers ─────────────────────────────────────────

    def _shdr(self, p: ctk.CTkFrame, text: str, row: int) -> None:
        """Section header: I N P U T ────────"""
        spaced = " ".join(text)  # approximate letter-spacing
        f = ctk.CTkFrame(p, fg_color="transparent")
        f.grid(row=row, column=0, padx=(18, 14), pady=(14, 3), sticky="ew")
        f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            f, text=spaced,
            text_color=FG_SUBTLE, anchor="w",
            font=ctk.CTkFont(family=_mono_font(), size=9),
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkFrame(
            f, height=1, fg_color=BORDER, corner_radius=0,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

    @staticmethod
    def _flbl(p: ctk.CTkFrame, text: str, row: int) -> None:
        ctk.CTkLabel(
            p, text=text,
            text_color=FG_MUTED, anchor="w",
            font=ctk.CTkFont(family=_prop_font(), size=11),
        ).grid(row=row, column=0, padx=18, pady=(5, 1), sticky="w")

    # ── Station list ──────────────────────────────────────────────────────────

    def _populate_stations(self) -> None:
        keys   = list_stations()
        values = [f"{STATIONS[k].name}  ({k})" for k in keys]
        self._station_keys = keys
        self._station_combo.configure(values=values)
        if values:
            self._station_combo.set(values[0])

    # ── Event handlers ────────────────────────────────────────────────────────

    def _on_browse(self) -> None:
        d = filedialog.askdirectory(initialdir=self._dir_var.get() or str(Path.cwd()))
        if d:
            self._output_modified = False
            self._dir_var.set(d)

    def _on_dir_typed(self, *_: Any) -> None:
        if hasattr(self, "_dir_debounce"):
            self.after_cancel(self._dir_debounce)
        self._dir_debounce = self.after(280, self._on_dir_settled)

    def _on_dir_settled(self) -> None:
        path = self._dir_var.get()
        self._auto_detect_station(path)
        self._suggest_filename(path)

    def _on_station_change(self, _: str = "") -> None:
        self._output_modified = False
        self._suggest_filename(self._dir_var.get())

    def _on_output_typed(self, *_: Any) -> None:
        if not self._auto_filling:
            self._output_modified = True

    def _auto_detect_station(self, path: str) -> None:
        pl = path.lower()
        for folder, alias in _folder_mapping().items():
            if folder.lower() in pl:
                st = get_station(alias)
                if st:
                    self._set_station_combo(alias)
                return

    def _set_station_combo(self, key: str) -> None:
        kl   = key.lower()
        vals = list(self._station_combo.cget("values"))
        for i, k in enumerate(self._station_keys):
            if k == kl:
                self._station_combo.set(vals[i]); return
        st = get_station(kl)
        if st:
            for i, k in enumerate(self._station_keys):
                if STATIONS[k].name == st.name:
                    self._station_combo.set(vals[i]); return

    def _suggest_filename(self, path: str) -> None:
        if self._output_modified:
            return
        p  = Path(path)
        q, yr, sf = p.name.lower(), p.parent.name, p.parent.parent.name.lower()
        if q.startswith("q") and yr.isdigit():
            self._auto_filling = True
            self._output_var.set(f"{yr}_{q}_{sf}.csv")
            self._auto_filling = False

    def _get_station_key(self) -> str | None:
        val  = self._station_var.get()
        vals = list(self._station_combo.cget("values"))
        if val in vals:
            idx = vals.index(val)
            if 0 <= idx < len(self._station_keys):
                return self._station_keys[idx]
        return None

    # ── Run ───────────────────────────────────────────────────────────────────

    def _on_run(self) -> None:
        if self._processing:
            return
        key = self._get_station_key()
        if not key:
            self._append_log("Error: No station selected.\n"); return
        wd = self._dir_var.get().strip()
        if not wd or not Path(wd).is_dir():
            self._append_log(f"Error: Directory not found: {wd}\n"); return
        outf = self._output_var.get().strip()
        if not outf:
            self._append_log("Error: Output filename required.\n"); return

        self._processing   = True
        self._run_btn.configure(state="disabled")
        self._set_status("running")
        self._orig_stdout  = sys.stdout
        sys.stdout         = LogWriter(self._log_q)

        threading.Thread(
            target=self._run_processing,
            args=(wd, key, outf,
                  self._no_sw_var.get(),
                  self._add_var.get().strip(),
                  self._no_rej_var.get()),
            daemon=True,
        ).start()

    def _run_processing(
        self, wd: str, key: str, outf: str,
        no_sw: bool, additional: str, no_rej: bool,
    ) -> None:
        orig_cwd = os.getcwd()
        handler  = GUIInteractionHandler(self)
        err      = False
        reject_path: Path | None = None
        try:
            os.chdir(wd)
            station = get_station(key)
            if station is None:
                print(f"Error: Unknown station '{key}'"); err = True; return

            outp = Path(wd) / outf
            if outp.exists():
                try:
                    with open(outp, "a"): pass
                    outp.unlink()
                    print(f"Overwriting existing {outf}")
                except (PermissionError, OSError):
                    print(f"Error: '{outf}' is locked — close it first.")
                    err = True; return

            if no_sw:
                print("Warning: Stopword filtering DISABLED")
            if station.convert:
                from main import run_convert_script
                run_convert_script()

            stats = ProcessingStats(output_file=outf)
            try:
                files = get_files(station, exclude_filename=outf)
            except SystemExit:
                err = True; return

            stats.files_total = len(files)
            content, rej_idx = read_files(
                files, station, stats, interaction_handler=handler)
            process_files(
                content=content, station=station, output_file=outp,
                additional_filter=additional, use_stopwords=not no_sw,
                stats=stats, force_reject_indices=rej_idx,
                save_reject_file=not no_rej,
            )
            print_summary_box(stats)

            if not no_rej:
                rp = generate_rejection_filename(station)
                if rp.exists() and rp.stat().st_size > 0:
                    reject_path = rp

        except SystemExit:
            pass
        except Exception as e:
            print(f"Error: {e}"); err = True
        finally:
            os.chdir(orig_cwd)
            self.after(0, lambda: self._on_done(err, reject_path))

    def _on_done(self, error: bool, reject_path: Path | None = None) -> None:
        sys.stdout       = self._orig_stdout
        self._processing = False
        self._run_btn.configure(state="normal")
        self._set_status("error" if error else "done")
        self._last_reject_path = reject_path
        if reject_path:
            self._open_rej_btn.grid()
        else:
            self._open_rej_btn.grid_remove()

    def _open_reject_file(self) -> None:
        if self._last_reject_path and self._last_reject_path.exists():
            import subprocess
            subprocess.Popen(["explorer", str(self._last_reject_path)])

    # ── Log ───────────────────────────────────────────────────────────────────

    def _tick_log(self) -> None:
        try:
            while True:
                self._append_log(self._log_q.get_nowait())
        except queue.Empty:
            pass
        self.after(self._LOG_MS, self._tick_log)

    def _append_log(self, text: str) -> None:
        self._log.configure(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _clear_log(self) -> None:
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ── Status & animations ───────────────────────────────────────────────────

    def _set_status(self, state: str) -> None:
        if state == "running":
            self._status_var.set("  \u25cf  Running\u2026")
            self._anim_tick = 0
            self._tick_anim()
        elif state == "done":
            self._status_var.set("  \u2713  Done")
            self._status_lbl.configure(text_color=C_SUCCESS)
        elif state == "error":
            self._status_var.set("  \u2717  Error")
            self._status_lbl.configure(text_color=C_ERROR)
        else:
            self._status_var.set("  \u25cf  Ready")
            self._status_lbl.configure(text_color=FG_SUBTLE)

    def _tick_anim(self) -> None:
        """Smooth sine-wave pulse shared by Run button and status dot."""
        if not self._processing:
            self._run_btn.configure(fg_color=ACCENT)
            return
        T     = 1400  # period ms
        phase = 2 * math.pi * self._anim_tick / T
        t_btn = (math.sin(phase) + 1) / 2               # 0 → 1
        t_dot = (math.sin(phase + math.pi) + 1) / 2     # offset by π
        self._run_btn.configure(fg_color=_lerp(ACCENT_DIM, ACCENT_HI, t_btn))
        self._status_lbl.configure(text_color=_lerp(ACCENT_DIM, ACCENT_HI, t_dot))
        self._anim_tick += self._ANIM_MS
        self.after(self._ANIM_MS, self._tick_anim)

    # ── Background animation ──────────────────────────────────────────────────

    def _tick_bg(self) -> None:
        if not _PIL:
            return
        try:
            w, h = self.winfo_width(), self.winfo_height()
            if w < 10 or h < 10:
                self.after(self._BG_MS, self._tick_bg); return
            if self._bg_rend is None:
                self._bg_rend = _BgRenderer(w, h)
            frame          = self._bg_rend.next_frame()
            self._bg_photo = PILPhotoImage(frame)
            self._bg_canvas.itemconfig(self._bg_item, image=self._bg_photo)
            self.after(self._BG_MS, self._tick_bg)
        except tk.TclError:
            pass  # window destroyed mid-frame

    def _on_resize(self, e: Any) -> None:
        if e.widget is not self:
            return
        wh = (e.width, e.height)
        if wh == self._last_wh:
            return
        self._last_wh = wh
        w, h = e.width, e.height

        if _PIL and self._bg_rend is not None:
            if hasattr(self, "_resize_job"):
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(
                80, lambda: self._bg_rend.resize(w, h))


# ── Entry point ───────────────────────────────────────────────────────────────
def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
