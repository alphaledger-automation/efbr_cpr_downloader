"""
eFBR CPR Downloader — GUI
==========================
Themeable GUI built with customtkinter.

Layout (top to bottom):
  1. Header            (title + Design Mode and Light/Dark switches)
  2. Source            (Excel file picker + CPR column selector)
  3. Controls          (Start / Pause-Resume / Cancel + status pill)
  4. Progress          (bar, percentage, Total/Success/Failed/Remaining)
  5. Activity log      (colour-coded, scrollable)
  6. Footer            (Open output folder / Open Excel summary)

When a run ends, a completion dialog shows the count per status and
offers the Excel summary and the output folder.

Appearance is driven entirely by themes.Palette: every widget registers a
"styler" callback, and switching Design Mode or Light/Dark just re-runs
them — nothing is rebuilt, so it is safe to switch during a download.
"""

import os
import queue
import platform
import subprocess
import threading
import tkinter as tk
from typing import Callable, Dict, List, Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox

from constants import (
    APP_NAME, APP_VERSION, get_output_dir,
    STATUS_DOWNLOADED, STATUS_NOT_EXIST, STATUS_LOGIN_ISSUE,
    STATUS_NOT_DOWNLOADED, load_output_prefs, save_output_prefs,
)
from excel_handler import get_column_names, detect_cpr_column, analyze_cprs
from report_writer import write_summary_report
from browser_engine import DownloadWorker, LogEntry, ProgressUpdate
from themes import (
    DESIGNS, MODES, Palette, blend, get_palette, load_prefs, save_prefs,
)


NEU_INSET = 5       # neumorphic: the painted card extends this far past its frame


def _gaps(p: Palette):
    """(left/right, above/below) gap around every card for a design."""
    return (28, 12) if p.design == "neumorphic" else (22, 7)


def _now() -> str:
    import datetime
    return datetime.datetime.now().strftime("%H:%M:%S")


def _open_path(path: str) -> None:
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class CPRDownloaderApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("960x770")
        self.minsize(860, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # ── Internal state ────────────────────────────────────────────────
        self._input_file: str = ""
        self._cpr_list: List[str] = []
        self._columns: List[str] = []
        self._worker: Optional[DownloadWorker] = None
        self._worker_thread: Optional[threading.Thread] = None
        self._stop_event  = threading.Event()
        self._pause_event = threading.Event()       # SET = paused
        self._log_q  = queue.Queue()
        self._prog_q = queue.Queue()
        self._is_running = False
        self._report_path = ""
        self._duplicates: List[dict] = []
        self._dup_extra = 0
        prefs = load_output_prefs()
        saved = prefs.get("output_dir")
        try:
            self._output_dir = str(saved) if saved else str(get_output_dir())
        except Exception:
            self._output_dir = str(get_output_dir())
        self._keep_name_var = tk.BooleanVar(value=prefs.get("keep_name", True))

        # ── Theme state ───────────────────────────────────────────────────
        self._design, self._mode = load_prefs()
        self._p: Palette = get_palette(self._design, self._mode)
        self._stylers: List[Callable[[Palette], None]] = []
        self._roles: Dict[ctk.CTkLabel, str] = {}       # label -> text role
        self._btn_kind: Dict[ctk.CTkButton, str] = {}   # button -> palette kind
        self._cards: List[ctk.CTkFrame] = []
        self._bg_job = None
        self._last_size = (0, 0)

        self._build_ui()
        self._apply_theme()
        self.bind("<Configure>", self._on_configure)
        self._poll_queues()

    # ==================================================================
    #  THEME PLUMBING
    # ==================================================================

    def _style(self, fn: Callable[[Palette], None]) -> None:
        """Register a callback that (re)colours one widget from a palette."""
        self._stylers.append(fn)

    def _card(self, row: int, sticky: str = "ew") -> ctk.CTkFrame:
        card = ctk.CTkFrame(self)
        card.grid(row=row, column=0, sticky=sticky)
        self._cards.append(card)

        def style_card(p: Palette, c=card, last=(row == 5)):
            gap_x, gap_y = _gaps(p)
            c.configure(
                fg_color=p.card, border_color=p.border,
                border_width=p.border_width, corner_radius=p.radius,
            )
            c.grid(padx=gap_x,           # .grid() applies DPI scaling
                   pady=(gap_y, gap_y + 8 if last else gap_y))
        self._style(style_card)
        return card

    def _row(self, parent, role: str = "card") -> ctk.CTkFrame:
        """Borderless layout frame that takes on its parent's surface colour."""
        frame = ctk.CTkFrame(parent, corner_radius=0)

        def style_row(p: Palette, f=frame):
            colour = p.header_bg if role == "header" else p.card
            f.configure(fg_color=colour, bg_color=colour)
        self._style(style_row)
        return frame

    def _label(self, parent, text: str = "", size: int = 12,
               bold: bool = False, role: str = "text", **kw) -> ctk.CTkLabel:
        lbl = ctk.CTkLabel(
            parent, text=text,
            font=ctk.CTkFont(size=size, weight="bold" if bold else "normal"),
            **kw,
        )
        self._roles[lbl] = role
        return lbl

    def _set(self, lbl: ctk.CTkLabel, text: Optional[str] = None,
             role: Optional[str] = None) -> None:
        """Update a label's text and/or semantic colour role."""
        if role is not None:
            self._roles[lbl] = role
        kw = {"text_color": self._p.color(self._roles[lbl])}
        if text is not None:
            kw["text"] = text
        lbl.configure(**kw)

    def _button(self, parent, text: str, kind: str, **kw) -> ctk.CTkButton:
        btn = ctk.CTkButton(parent, text=text, **kw)
        self._btn_kind[btn] = kind
        return btn

    def _button_colors(self, kind: str, disabled: bool = False) -> dict:
        """configure() kwargs that paint a button of the given palette kind."""
        p = self._p
        fill, hover, text, border_w = p.buttons[kind]
        if disabled:
            fill = blend(fill, p.card, 0.6)
        return dict(
            fg_color=fill, hover_color=hover, text_color=text,
            text_color_disabled=blend(fill, text, 0.45),
            border_width=border_w, border_color=p.input_border,
            corner_radius=6 if p.design == "classic" else 10,
        )

    def _paint_button(self, btn: ctk.CTkButton) -> None:
        btn.configure(**self._button_colors(
            self._btn_kind[btn], btn.cget("state") == "disabled"))

    def _set_state(self, btn: ctk.CTkButton, state: str) -> None:
        btn.configure(state=state)
        self._paint_button(btn)

    def _apply_theme(self) -> None:
        p = self._p = get_palette(self._design, self._mode)
        ctk.set_appearance_mode(self._mode)         # title bar, scrollbars
        self.configure(fg_color=p.bg_bottom)
        for fn in self._stylers:
            fn(p)
        for lbl in self._roles:
            self._set(lbl)
        for btn in self._btn_kind:
            self._paint_button(btn)
        self.update_idletasks()
        self._redraw_backdrop()

    def _on_design_change(self, choice: str) -> None:
        self._design = DESIGNS[choice]
        save_prefs(self._design, self._mode)
        self._apply_theme()

    def _on_mode_change(self, choice: str) -> None:
        self._mode = MODES[choice]
        save_prefs(self._design, self._mode)
        self._apply_theme()

    # ── Backdrop (gradient + neumorphic shadows, painted behind the cards) ─

    def _on_configure(self, event) -> None:
        if event.widget is not self:
            return
        size = (event.width, event.height)
        if size == self._last_size:
            return
        self._last_size = size
        if self._bg_job is not None:
            self.after_cancel(self._bg_job)
        self._bg_job = self.after(40, self._redraw_backdrop)

    @staticmethod
    def _round_rect(canvas: tk.Canvas, x1, y1, x2, y2, r, fill: str) -> None:
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        canvas.create_polygon(pts, smooth=True, fill=fill, outline="")

    def _redraw_backdrop(self) -> None:
        self._bg_job = None
        p, c = self._p, self._backdrop
        w, h = self.winfo_width(), self.winfo_height()
        try:
            scale = self._get_window_scaling()
        except Exception:
            scale = 1.0

        c.delete("all")
        c.configure(bg=p.bg_bottom)

        # Vertical gradient (glass)
        if p.bg_top != p.bg_bottom and h > 1:
            band = 2
            for y in range(0, h, band):
                c.create_rectangle(
                    0, y, w, y + band, width=0,
                    fill=blend(p.bg_top, p.bg_bottom, y / h),
                )

        # Neumorphic: the visible card is painted here, slightly larger than
        # its frame, so the rounded corners and both soft shadows show.
        # Dark shadow falls bottom-right, light shadow top-left.
        if p.shadow_dark:
            inset, off, blur = NEU_INSET * scale, 4 * scale, 5 * scale
            r, steps = p.radius * scale, 6
            boxes = [
                (k.winfo_x() - inset, k.winfo_y() - inset,
                 k.winfo_x() + k.winfo_width() + inset,
                 k.winfo_y() + k.winfo_height() + inset)
                for k in self._cards
            ]
            for sign, colour in ((1, p.shadow_dark), (-1, p.shadow_light)):
                d = sign * off
                for x1, y1, x2, y2 in boxes:
                    for i in range(steps):
                        grow = blur * (1 - i / (steps - 1))
                        self._round_rect(
                            c, x1 + d - grow, y1 + d - grow,
                            x2 + d + grow, y2 + d + grow, r + grow,
                            blend(p.bg_bottom, colour, ((i + 1) / steps) ** 2),
                        )
            for x1, y1, x2, y2 in boxes:
                self._round_rect(c, x1, y1, x2, y2, r, p.card)

        # Card corners sit on the backdrop — match its colour there
        for card in self._cards:
            mid_y = card.winfo_y() + card.winfo_height() / 2
            corner = blend(p.bg_top, p.bg_bottom, mid_y / max(h, 1))
            if card.cget("bg_color") != corner:
                card.configure(bg_color=corner)

    # ==================================================================
    #  UI BUILD
    # ==================================================================

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)         # log area stretches

        self._backdrop = tk.Canvas(self, highlightthickness=0, bd=0)
        self._backdrop.place(x=0, y=0, relwidth=1, relheight=1)

        self._build_header(row=0)
        self._build_source(row=1)
        self._build_controls(row=2)
        self._build_progress(row=3)
        self._build_log(row=4)
        self._build_footer(row=5)

    # ── 1. Header ─────────────────────────────────────────────────────────

    def _build_header(self, row: int):
        header = self._card(row)
        header.grid_columnconfigure(0, weight=1)

        def style_header(p: Palette):
            classic = p.design == "classic"
            header.configure(
                fg_color=p.header_bg,
                corner_radius=0 if classic else p.radius,
            )
            gap_x, gap_y = _gaps(p)
            header.grid(
                padx=0 if classic else gap_x,
                pady=(0, gap_y) if classic else (gap_y + 8, gap_y),
            )
        self._style(style_header)

        titles = self._row(header, "header")
        titles.grid(row=0, column=0, sticky="w", padx=18, pady=10)
        self._label(titles, APP_NAME, size=18, bold=True,
                    role="h_text").pack(anchor="w")
        self._label(
            titles, f"Bulk-download CPR receipts from e.fbr.gov.pk  ·  v{APP_VERSION}",
            size=11, role="h_muted",
        ).pack(anchor="w")

        switches = self._row(header, "header")
        switches.grid(row=0, column=1, sticky="e", padx=18, pady=10)

        def segmented(col: int, caption: str, values, current: str, command):
            self._label(switches, caption, size=11, role="h_muted").grid(
                row=0, column=col, sticky="w", padx=(14, 0))
            seg = ctk.CTkSegmentedButton(
                switches, values=list(values), command=command,
                font=ctk.CTkFont(size=12), height=30,
            )
            seg.set(current)
            seg.grid(row=1, column=col, padx=(14, 0), pady=(2, 0))

            def style_segmented(p: Palette, s=seg):
                s.configure(
                    bg_color=p.header_bg,
                    fg_color=p.seg_bg, unselected_color=p.seg_bg,
                    unselected_hover_color=p.seg_hover,
                    selected_color=p.seg_selected,
                    selected_hover_color=p.seg_selected,
                    text_color=p.header_text,
                    corner_radius=6 if p.design == "classic" else 10,
                )
                # customtkinter leaves the inner buttons' backing colour
                # stale when fg_color changes, which shows as 1px edges.
                for inner in getattr(s, "_buttons_dict", {}).values():
                    inner.configure(bg_color=p.seg_bg)
            self._style(style_segmented)
            return seg

        design_name = next(k for k, v in DESIGNS.items() if v == self._design)
        mode_name = next(k for k, v in MODES.items() if v == self._mode)
        self._seg_design = segmented(
            0, "Design Mode", DESIGNS, design_name, self._on_design_change)
        self._seg_mode = segmented(
            1, "Appearance", MODES, mode_name, self._on_mode_change)

    # ── 2. Source (file + column) ─────────────────────────────────────────

    def _build_source(self, row: int):
        card = self._card(row)
        card.grid_columnconfigure(1, weight=1)

        self._label(card, "Excel File", size=13, bold=True).grid(
            row=0, column=0, sticky="w", padx=(18, 12), pady=(12, 5))

        file_row = self._row(card)
        file_row.grid(row=0, column=1, sticky="ew", padx=(0, 18), pady=(12, 5))
        file_row.grid_columnconfigure(1, weight=1)

        self._button(
            file_row, "Browse…", "primary", width=110, height=32,
            command=self._browse_file,
        ).grid(row=0, column=0, padx=(0, 12))

        self._file_label = self._label(
            file_row, "No file selected", role="muted", anchor="w")
        self._file_label.grid(row=0, column=1, sticky="ew")

        self._records_label = self._label(file_row, "", bold=True, role="ok")
        self._records_label.grid(row=0, column=2, padx=(12, 0))

        self._label(card, "CPR Column", size=13, bold=True).grid(
            row=1, column=0, sticky="w", padx=(18, 12), pady=(5, 5))

        col_row = self._row(card)
        col_row.grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=(5, 5))
        col_row.grid_columnconfigure(1, weight=1)

        self._col_dropdown = ctk.CTkComboBox(
            col_row, values=["(auto-detect)"], width=300, height=32,
            state="readonly", command=self._on_column_change,
        )
        self._col_dropdown.set("(auto-detect)")
        self._col_dropdown.grid(row=0, column=0, sticky="w")
        self._style(lambda p: self._col_dropdown.configure(
            fg_color=p.input_bg, border_color=p.input_border,
            button_color=p.input_border, button_hover_color=p.accent,
            text_color=p.text, dropdown_fg_color=p.card,
            dropdown_text_color=p.text, dropdown_hover_color=p.hover,
            corner_radius=6 if p.design == "classic" else 10,
        ))

        self._col_status = self._label(col_row, "", size=11, role="muted",
                                       anchor="e")
        self._col_status.grid(row=0, column=1, sticky="e", padx=(12, 0))

        # Save-to folder
        self._label(card, "Save To", size=13, bold=True).grid(
            row=2, column=0, sticky="w", padx=(18, 12), pady=(5, 5))
        out_row = self._row(card)
        out_row.grid(row=2, column=1, sticky="ew", padx=(0, 18), pady=(5, 5))
        out_row.grid_columnconfigure(1, weight=1)
        self._button(
            out_row, "Choose Folder…", "primary", width=110, height=32,
            command=self._choose_output_dir,
        ).grid(row=0, column=0, padx=(0, 12))
        self._out_label = self._label(out_row, self._output_dir, role="muted",
                                      anchor="w")
        self._out_label.grid(row=0, column=1, sticky="ew")

        # File naming option
        self._label(card, "File Name", size=13, bold=True).grid(
            row=3, column=0, sticky="w", padx=(18, 12), pady=(5, 12))
        name_row = self._row(card)
        name_row.grid(row=3, column=1, sticky="ew", padx=(0, 18), pady=(5, 12))
        self._chk_keep = ctk.CTkCheckBox(
            name_row, text="Keep original file name (do not rename to CPR_<no>.pdf)",
            variable=self._keep_name_var, command=self._persist_output_prefs,
        )
        self._chk_keep.grid(row=0, column=0, sticky="w")
        self._style(lambda p: self._chk_keep.configure(
            text_color=p.text, fg_color=p.accent, hover_color=p.accent,
            border_color=p.input_border))

    # ── 3. Controls ───────────────────────────────────────────────────────

    def _build_controls(self, row: int):
        card = self._card(row)
        card.grid_columnconfigure(3, weight=1)
        big = ctk.CTkFont(size=14, weight="bold")

        self._btn_start = self._button(
            card, "Start Download", "start", font=big, width=170, height=40,
            command=self._start_download,
        )
        self._btn_start.grid(row=0, column=0, padx=(18, 10), pady=11)

        self._btn_pause = self._button(
            card, "Pause", "pause", font=big, width=120, height=40,
            state="disabled", command=self._toggle_pause,
        )
        self._btn_pause.grid(row=0, column=1, padx=(0, 10), pady=11)

        self._btn_cancel = self._button(
            card, "Cancel", "cancel", font=big, width=120, height=40,
            state="disabled", command=self._cancel_download,
        )
        self._btn_cancel.grid(row=0, column=2, pady=11)

        self._status_pill = self._label(
            card, "●  Ready", size=12, bold=True, role="muted",
            corner_radius=14, height=28, width=120,
        )
        self._status_pill.grid(row=0, column=3, sticky="e", padx=18)
        self._style(lambda p: self._status_pill.configure(fg_color=p.input_bg))

    # ── 4. Progress ───────────────────────────────────────────────────────

    def _build_progress(self, row: int):
        card = self._card(row)
        card.grid_columnconfigure(0, weight=1)

        top = self._row(card)
        top.grid(row=0, column=0, sticky="ew", padx=18, pady=(11, 5))
        top.grid_columnconfigure(1, weight=1)

        self._lbl_progress = self._label(top, "0 / 0 processed", size=13,
                                         bold=True)
        self._lbl_progress.grid(row=0, column=0, sticky="w")

        self._lbl_current = self._label(top, "", role="muted", anchor="e")
        self._lbl_current.grid(row=0, column=1, sticky="e", padx=12)

        self._lbl_percent = self._label(top, "0%", size=13, bold=True,
                                        role="accent")
        self._lbl_percent.grid(row=0, column=2, sticky="e")

        self._progress_bar = ctk.CTkProgressBar(card, height=12)
        self._progress_bar.grid(row=1, column=0, sticky="ew", padx=18)
        self._progress_bar.set(0)
        self._style(lambda p: self._progress_bar.configure(
            fg_color=p.track, progress_color=p.accent,
            corner_radius=3 if p.design == "classic" else 6,
        ))

        stats = self._row(card)
        stats.grid(row=2, column=0, sticky="ew", padx=12, pady=(10, 12))

        def tile(col: int, caption: str, role: str) -> ctk.CTkLabel:
            stats.grid_columnconfigure(col, weight=1, uniform="stat")
            box = ctk.CTkFrame(stats)
            box.grid(row=0, column=col, sticky="ew", padx=6)
            self._style(lambda p, b=box: b.configure(
                fg_color=p.input_bg,
                corner_radius=6 if p.design == "classic" else 12,
            ))
            value = self._label(box, "0", size=18, bold=True, role=role)
            value.pack(anchor="w", padx=14, pady=(5, 0))
            self._label(box, caption, size=11, role="muted").pack(
                anchor="w", padx=14, pady=(0, 5))
            return value

        self._stat_total   = tile(0, "Unique CPRs", "text")
        self._stat_success = tile(1, "Downloaded", "ok")
        self._stat_failed  = tile(2, "Failed", "err")
        self._stat_left    = tile(3, "Remaining", "text")
        self._stat_dup     = tile(4, "Duplicates (skipped)", "warn")

    # ── 5. Log ────────────────────────────────────────────────────────────

    def _build_log(self, row: int):
        card = self._card(row, sticky="nsew")
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        hdr = self._row(card)
        hdr.grid(row=0, column=0, sticky="ew", padx=18, pady=(11, 6))
        self._label(hdr, "Activity Log", size=13, bold=True).pack(side="left")
        self._button(
            hdr, "Clear", "secondary", width=70, height=26,
            command=self._clear_log,
        ).pack(side="right")

        self._log_box = ctk.CTkTextbox(
            card, font=ctk.CTkFont(family="Consolas", size=12),
            state="disabled", wrap="word", height=120,
        )
        self._log_box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 13))

        def style_log(p: Palette):
            self._log_box.configure(
                fg_color=p.log_bg, text_color=p.text,
                border_width=1, border_color=p.input_border,
                scrollbar_button_color=p.input_border,
                scrollbar_button_hover_color=p.accent,
                corner_radius=6 if p.design == "classic" else 12,
            )
            # Colour tags for the log levels
            self._log_box.tag_config("ok",   foreground=p.ok)
            self._log_box.tag_config("err",  foreground=p.err)
            self._log_box.tag_config("warn", foreground=p.warn)
            self._log_box.tag_config("info", foreground=p.info)
        self._style(style_log)

    # ── 6. Footer ─────────────────────────────────────────────────────────

    def _build_footer(self, row: int):
        card = self._card(row)
        card.grid_columnconfigure(2, weight=1)

        self._btn_folder = self._button(
            card, "Open Output Folder", "secondary", width=190, height=34,
            command=self._open_downloads_folder,
        )
        self._btn_folder.grid(row=0, column=0, padx=(18, 10), pady=10)

        self._btn_report = self._button(
            card, "Open Excel Summary", "secondary", width=190, height=34,
            state="disabled", command=self._open_report,
        )
        self._btn_report.grid(row=0, column=1, pady=10)

        self._foot_label = self._label(
            card, f"Saving to  {self._output_dir}", size=11, role="muted",
            anchor="e",
        )
        self._foot_label.grid(row=0, column=2, sticky="e", padx=18)

    # ==================================================================
    #  EVENT HANDLERS
    # ==================================================================

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File with CPR Numbers",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not path:
            return

        self._input_file = path
        self._set(self._file_label, os.path.basename(path), "text")

        try:
            self._columns = get_column_names(path)
            self._col_dropdown.configure(
                values=["(auto-detect)"] + self._columns,
            )
            self._col_dropdown.set("(auto-detect)")
            self._load_cprs(column_name=None)
        except Exception as exc:
            messagebox.showerror("File Error", str(exc))
            self._set(self._records_label, "")
            self._set(self._col_status, "")

    def _choose_output_dir(self):
        if self._is_running:
            messagebox.showinfo("Busy", "Change the folder after the run ends.")
            return
        path = filedialog.askdirectory(
            title="Choose folder to save CPR files",
            initialdir=self._output_dir if os.path.isdir(self._output_dir)
            else None,
            mustexist=False,
        )
        if not path:
            return
        self._output_dir = os.path.normpath(path)
        self._set(self._out_label, self._output_dir)
        self._set(self._foot_label, f"Saving to  {self._output_dir}")
        self._persist_output_prefs()

    def _persist_output_prefs(self):
        save_output_prefs(self._output_dir, bool(self._keep_name_var.get()))

    def _out_path(self):
        import pathlib
        p = pathlib.Path(self._output_dir)
        p.mkdir(parents=True, exist_ok=True)
        return p

    def _on_column_change(self, choice: str):
        if not self._input_file:
            return
        col = None if choice == "(auto-detect)" else choice
        self._load_cprs(column_name=col)

    def _load_cprs(self, column_name: Optional[str]):
        """Read CPRs from the selected file+column and update the UI."""
        try:
            info = analyze_cprs(self._input_file, column_name)
            self._cpr_list = info["unique"]
            self._duplicates = info["duplicates"]
            self._dup_extra = info["dup_extra"]
            label = f"{len(self._cpr_list):,} unique CPRs"
            if self._dup_extra:
                label += f"  ({self._dup_extra:,} duplicate rows)"
            self._set(self._records_label, label,
                      "warn" if self._dup_extra else "ok")
            if self._dup_extra:
                self._append_log(
                    f"[{_now()}] {len(self._duplicates)} duplicate CPR(s) "
                    f"found ({self._dup_extra} extra rows) - each will be "
                    "downloaded only once:", "warn")
                for d in self._duplicates:
                    rows = ", ".join(str(r) for r in d["rows"])
                    self._append_log(
                        f"    {d['cpr']}  x{d['count']}  (Excel rows {rows})",
                        "warn")
            if not self._is_running:
                self._set_stats(len(self._cpr_list), 0, 0)
            if column_name:
                self._set(self._col_status,
                          f'Using column: "{column_name}"', "muted")
            else:
                detected = detect_cpr_column(self._input_file)
                if detected:
                    self._set(self._col_status,
                              f'Auto-detected: "{detected}"', "ok")
                    self._col_dropdown.set(detected)
                else:
                    self._set(self._col_status,
                              "Auto-detect failed — select manually", "err")
        except Exception as exc:
            self._cpr_list = []
            self._duplicates = []
            self._dup_extra = 0
            self._set(self._records_label, "")
            self._set(self._col_status, str(exc), "err")
            self._set_stats(0, 0, 0)

    def _set_status(self, text: str, role: str):
        self._set(self._status_pill, f"●  {text}", role)

    def _set_stats(self, total: int, success: int, failed: int):
        self._set(self._stat_total, f"{total:,}")
        self._set(self._stat_success, f"{success:,}")
        self._set(self._stat_failed, f"{failed:,}")
        self._set(self._stat_left, f"{max(total - success - failed, 0):,}")
        self._set(self._stat_dup, f"{self._dup_extra:,}")

    def _refresh_stats(self):
        """Recount successes/failures from the worker's results so far."""
        if not self._worker:
            return
        results = list(self._worker.results)
        success = sum(1 for r in results if r.get("Status") == STATUS_DOWNLOADED)
        self._set_stats(len(self._worker.cpr_list), success,
                        len(results) - success)

    # ── Start / Pause / Cancel ────────────────────────────────────────────

    def _start_download(self):
        if self._is_running:
            return
        if not self._input_file:
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return
        if not self._cpr_list:
            messagebox.showwarning(
                "No CPRs",
                "No CPR numbers found. Check your file and column selection.",
            )
            return

        self._stop_event.clear()
        self._pause_event.clear()
        self._is_running = True
        self._progress_bar.set(0)
        self._set(self._lbl_progress, f"0 / {len(self._cpr_list)} processed")
        self._set(self._lbl_percent, "0%")
        self._set(self._lbl_current, "")
        self._set_stats(len(self._cpr_list), 0, 0)
        self._set_status("Running", "accent")

        self._set_state(self._btn_start, "disabled")
        self._btn_pause.configure(text="Pause")
        self._set_state(self._btn_pause, "normal")
        self._set_state(self._btn_cancel, "normal")
        self._set_state(self._btn_report, "disabled")

        try:
            output_dir = self._out_path()
        except Exception as exc:
            messagebox.showerror("Folder Error",
                                 f"Cannot use save folder:\n{exc}")
            self._is_running = False
            self._set_state(self._btn_start, "normal")
            self._set_state(self._btn_pause, "disabled")
            self._set_state(self._btn_cancel, "disabled")
            self._set_status("Ready", "muted")
            return
        self._persist_output_prefs()

        self._worker = DownloadWorker(
            cpr_list=list(self._cpr_list),
            output_dir=output_dir,
            log_q=self._log_q,
            prog_q=self._prog_q,
            stop_event=self._stop_event,
            pause_event=self._pause_event,
            keep_original_name=bool(self._keep_name_var.get()),
            duplicates=list(self._duplicates),
        )
        self._worker_thread = threading.Thread(
            target=self._worker.run, daemon=True,
        )
        self._worker_thread.start()
        self._append_log(
            f"[{_now()}] Download started — {len(self._cpr_list)} unique CPRs"
            + (f" ({self._dup_extra} duplicate rows skipped)"
               if self._dup_extra else ""),
            "info",
        )


    def _toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self._btn_pause.configure(text="Pause")
            self._set_status("Running", "accent")
            self._append_log(f"[{_now()}] Resumed.", "info")
        else:
            self._pause_event.set()
            self._btn_pause.configure(text="Resume")
            self._set_status("Paused", "warn")
            self._append_log(
                f"[{_now()}] Paused — click Resume to continue.", "warn",
            )

    def _cancel_download(self):
        self._stop_event.set()
        self._pause_event.clear()
        self._set_state(self._btn_cancel, "disabled")
        self._set_state(self._btn_pause, "disabled")
        self._set_status("Cancelling…", "warn")
        self._append_log(
            f"[{_now()}] Cancel requested — finishing current CPR...", "warn",
        )

    # ── Footer actions ────────────────────────────────────────────────────

    def _open_downloads_folder(self):
        _open_path(self._output_dir)

    def _open_report(self):
        if self._report_path and os.path.exists(self._report_path):
            _open_path(self._report_path)

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _on_close(self):
        if self._is_running:
            if not messagebox.askyesno(
                "Quit", "Download is in progress. Stop and quit?",
            ):
                return
            self._stop_event.set()
            self._pause_event.clear()
        self.destroy()

    # ==================================================================
    #  QUEUE POLLING  (runs on the GUI / main thread every 150 ms)
    # ==================================================================

    def _poll_queues(self):
        # ── Log queue ─────────────────────────────────────────────────────
        for _ in range(50):
            try:
                entry: LogEntry = self._log_q.get_nowait()
                self._append_log(
                    f"[{entry.timestamp}] {entry.message}", entry.level,
                )
            except queue.Empty:
                break

        # ── Progress queue ────────────────────────────────────────────────
        for _ in range(10):
            try:
                upd: ProgressUpdate = self._prog_q.get_nowait()
                if upd.is_done:
                    self._on_done(upd)
                else:
                    self._update_progress(upd)
            except queue.Empty:
                break

        self.after(150, self._poll_queues)

    def _append_log(self, text: str, tag: str = "info"):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", text + "\n", tag)
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _update_progress(self, upd: ProgressUpdate):
        self._set(self._lbl_progress,
                  f"{upd.processed} / {upd.total} processed")
        if upd.current_cpr:
            self._set(self._lbl_current, f"Current: {upd.current_cpr}")
        if upd.total > 0:
            fraction = upd.processed / upd.total
            self._progress_bar.set(fraction)
            self._set(self._lbl_percent, f"{fraction:.0%}")
        self._refresh_stats()

    def _on_done(self, upd: ProgressUpdate):
        self._is_running = False
        cancelled = self._stop_event.is_set()
        self._set_state(self._btn_start, "normal")
        self._btn_pause.configure(text="Pause")
        self._set_state(self._btn_pause, "disabled")
        self._set_state(self._btn_cancel, "disabled")
        if not cancelled:
            self._progress_bar.set(1.0)
            self._set(self._lbl_percent, "100%")
        self._set(self._lbl_progress,
                  f"{upd.processed} / {upd.total} processed — "
                  + ("Cancelled" if cancelled else "Done"))
        self._set(self._lbl_current, "")
        self._refresh_stats()
        if cancelled:
            self._set_status("Cancelled", "warn")
        else:
            self._set_status("Completed", "ok")

        # Excel summary: one row per CPR, including any never attempted
        rows = self._worker.full_results()
        try:
            self._report_path = write_summary_report(
                rows, self._out_path(), self._duplicates)
            self._append_log(
                f"[{_now()}] Excel summary saved: {self._report_path}", "ok")
            self._set_state(self._btn_report, "normal")
        except Exception as exc:
            self._report_path = ""
            self._append_log(
                f"[{_now()}] Failed to write Excel summary: {exc}", "err")

        self._show_finish_dialog(rows, cancelled)

    def _show_finish_dialog(self, rows: List[dict], cancelled: bool):
        """Completion dialog: count per status + Excel summary / output folder."""
        p = self._p
        dlg = ctk.CTkToplevel(self, fg_color=p.card)
        dlg.title("Download cancelled" if cancelled else "Download complete")
        dlg.resizable(False, False)
        dlg.transient(self)

        body = ctk.CTkFrame(dlg, fg_color=p.card, corner_radius=0)
        body.pack(padx=26, pady=22)

        ctk.CTkLabel(
            body, text="Download cancelled" if cancelled else "Download complete",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=p.warn if cancelled else p.ok,
        ).pack(anchor="w")
        ctk.CTkLabel(
            body, text=f"{len(rows):,} unique CPRs in this run"
            + (f"  ·  {self._dup_extra:,} duplicate rows skipped"
               if self._dup_extra else ""),
            font=ctk.CTkFont(size=12), text_color=p.muted,
        ).pack(anchor="w", pady=(0, 12))

        table = ctk.CTkFrame(
            body, fg_color=p.input_bg,
            corner_radius=6 if p.design == "classic" else 12,
        )
        table.pack(fill="x")
        table.grid_columnconfigure(0, weight=1)
        for i, (status, colour) in enumerate((
            (STATUS_DOWNLOADED, p.ok),
            (STATUS_NOT_EXIST, p.warn),
            (STATUS_LOGIN_ISSUE, p.err),
            (STATUS_NOT_DOWNLOADED, p.muted),
        )):
            count = sum(1 for r in rows if r.get("Status") == status)
            pady = (10 if i == 0 else 3, 10 if i == 3 else 3)
            ctk.CTkLabel(
                table, text=status, font=ctk.CTkFont(size=13),
                text_color=p.text, anchor="w",
            ).grid(row=i, column=0, sticky="w", padx=(16, 60), pady=pady)
            ctk.CTkLabel(
                table, text=f"{count:,}",
                font=ctk.CTkFont(size=14, weight="bold"), text_color=colour,
            ).grid(row=i, column=1, sticky="e", padx=16, pady=pady)

        ctk.CTkLabel(
            body, text=f"Output folder:  {self._output_dir}",
            font=ctk.CTkFont(size=11), text_color=p.muted,
        ).pack(anchor="w", pady=(10, 14))

        actions = ctk.CTkFrame(body, fg_color=p.card, corner_radius=0)
        actions.pack(fill="x")
        has_report = bool(self._report_path)
        ctk.CTkButton(
            actions, text="Open Excel Summary", width=170, height=36,
            state="normal" if has_report else "disabled",
            command=self._open_report,
            **self._button_colors("primary", disabled=not has_report),
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            actions, text="Open Output Folder", width=170, height=36,
            command=self._open_downloads_folder,
            **self._button_colors("secondary"),
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            actions, text="Close", width=90, height=36, command=dlg.destroy,
            **self._button_colors("secondary"),
        ).pack(side="left")

        # Centre over the main window, then make it modal
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_reqheight()) // 2
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.after(150, lambda: (dlg.lift(), dlg.focus_force(), dlg.grab_set()))
        self._finish_dialog = dlg
