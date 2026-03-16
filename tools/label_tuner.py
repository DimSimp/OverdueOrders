#!/usr/bin/env python
"""Label Print Tuner — interactively adjust scale and split ratio per courier.

Run from the project root:
    python tools/label_tuner.py
"""
from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox

# Ensure project root is on the path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# ── Paths (relative to project root, which is the cwd when run normally) ────
LABELS_DIR = Path("data/shipping")
SETTINGS_FILE = LABELS_DIR / "label_settings.json"
CAPTURED_FILE = LABELS_DIR / "captured.json"

COURIER_DISPLAY = {
    "dai_post": "DAI Post",
    "aramex": "Aramex",
    "auspost": "Australia Post (Standard)",
    "auspost_express": "Australia Post (Express)",
    "allied": "Allied Express",
    "bonds": "Bonds",
}

DEFAULTS = {
    "scale": 0.92,
    "split_ratio": 0.57,
    "label_length_mm": 0.0,
    "no_split": False,
    "rotate_cw": False,
    "skip_initial_rotate": False,
}


def _load_settings(courier_code: str) -> dict:
    try:
        from src.shipping.label_settings import COURIER_DEFAULTS
    except Exception:
        COURIER_DEFAULTS = {}
    base = {**DEFAULTS, **COURIER_DEFAULTS.get(courier_code, {})}
    if SETTINGS_FILE.exists():
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if courier_code in data:
                return {**base, **data[courier_code]}
        except Exception:
            pass
    return base


def _save_settings(
    courier_code: str,
    scale: float,
    split_ratio: float,
    label_length_mm: float = 0.0,
    no_split: bool = False,
    rotate_cw: bool = False,
    skip_initial_rotate: bool = False,
) -> None:
    LABELS_DIR.mkdir(parents=True, exist_ok=True)
    existing: dict = {}
    if SETTINGS_FILE.exists():
        try:
            existing = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    existing[courier_code] = {
        "scale": round(scale, 4),
        "split_ratio": round(split_ratio, 4),
        "label_length_mm": round(label_length_mm, 1),
        "no_split": no_split,
        "rotate_cw": rotate_cw,
        "skip_initial_rotate": skip_initial_rotate,
    }
    SETTINGS_FILE.write_text(json.dumps(existing, indent=2), encoding="utf-8")


class LabelTuner(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Label Print Tuner")
        self.resizable(True, True)
        self._pdf_bytes: bytes | None = None
        self._courier_code: str = ""
        self._preview_photos: list = []   # prevent GC
        self._after_id: str | None = None  # debounce preview updates
        self._page_num: int = 0
        self._page_count: int = 1
        self._strips_cache: list | None = None  # cached process_label_pdf output
        self._build_ui()
        self._refresh_courier_list()

    # ── UI construction ────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=12, pady=12)

        # ── Left: controls ──────────────────────────────────────────────
        left = tk.LabelFrame(main, text="Settings", padx=10, pady=10)
        left.pack(side="left", fill="y", padx=(0, 12))

        tk.Label(left, text="Courier:").pack(anchor="w")
        self._courier_var = tk.StringVar()
        self._courier_combo = ttk.Combobox(
            left, textvariable=self._courier_var, width=24, state="readonly"
        )
        self._courier_combo.pack(fill="x", pady=(2, 12))
        self._courier_combo.bind("<<ComboboxSelected>>", self._on_courier_changed)

        tk.Label(left, text="Scale (%):").pack(anchor="w")
        self._scale_var = tk.DoubleVar(value=92)
        scale_row = tk.Frame(left)
        scale_row.pack(fill="x", pady=(2, 8))
        tk.Scale(
            scale_row, from_=50, to=250, resolution=1, orient="horizontal",
            variable=self._scale_var, command=self._on_param_changed, length=200,
        ).pack(side="left")
        tk.Spinbox(
            scale_row, from_=50, to=250, increment=1,
            textvariable=self._scale_var, width=5,
            command=self._on_param_changed,
        ).pack(side="left", padx=(6, 0))

        tk.Label(left, text="Split position (%):").pack(anchor="w", pady=(4, 0))
        self._split_var = tk.DoubleVar(value=57)
        split_row = tk.Frame(left)
        split_row.pack(fill="x", pady=(2, 4))
        self._split_scale = tk.Scale(
            split_row, from_=40, to=70, resolution=1, orient="horizontal",
            variable=self._split_var, command=self._on_param_changed, length=200,
        )
        self._split_scale.pack(side="left")
        self._split_spin = tk.Spinbox(
            split_row, from_=40, to=70, increment=1,
            textvariable=self._split_var, width=5,
            command=self._on_param_changed,
        )
        self._split_spin.pack(side="left", padx=(6, 0))

        self._nosplit_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            left, text="No split (single strip)",
            variable=self._nosplit_var, command=self._on_nosplit_toggled,
        ).pack(anchor="w", pady=(4, 4))

        self._rotate_cw_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            left, text="Rotate strips CW (Allied-style)",
            variable=self._rotate_cw_var, command=self._on_param_changed,
        ).pack(anchor="w", pady=(0, 2))

        self._skip_initial_rotate_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            left, text="Skip initial rotation (landscape labels e.g. Allied)",
            variable=self._skip_initial_rotate_var, command=self._on_param_changed,
            wraplength=220, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(left, text="Label length (mm):").pack(anchor="w", pady=(6, 0))
        self._length_auto_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            left, text="Auto (natural height)",
            variable=self._length_auto_var, command=self._on_length_auto_toggled,
        ).pack(anchor="w", pady=(2, 2))
        self._length_var = tk.DoubleVar(value=80)
        length_row = tk.Frame(left)
        length_row.pack(fill="x", pady=(0, 8))
        self._length_scale = tk.Scale(
            length_row, from_=30, to=300, resolution=5, orient="horizontal",
            variable=self._length_var, command=self._on_param_changed, length=200,
            state="disabled",
        )
        self._length_scale.pack(side="left")
        self._length_spin = tk.Spinbox(
            length_row, from_=30, to=300, increment=5,
            textvariable=self._length_var, width=5,
            command=self._on_param_changed, state="disabled",
        )
        self._length_spin.pack(side="left", padx=(6, 0))

        tk.Button(
            left, text="Print Test", width=18,
            bg="#1a5faa", fg="white", activebackground="#124080",
            command=self._print_test,
        ).pack(pady=(4, 6))

        tk.Button(
            left, text="Save Settings", width=18,
            bg="#1a7a1a", fg="white", activebackground="#115511",
            command=self._save_settings_ui,
        ).pack(pady=(0, 8))

        self._status_lbl = tk.Label(
            left, text="", wraplength=240, justify="left",
            font=("Segoe UI", 9), fg="gray40",
        )
        self._status_lbl.pack(anchor="w")

        # ── Right: preview ──────────────────────────────────────────────
        right = tk.LabelFrame(main, text="Preview  (as printed on tape)", padx=10, pady=10)
        right.pack(side="left", fill="both", expand=True)

        hdr = tk.Frame(right)
        hdr.pack(fill="x")
        self._hdr_top_lbl = tk.Label(hdr, text="Strip 1 — top half", font=("Segoe UI", 9, "bold"))
        self._hdr_top_lbl.pack(side="left", padx=(0, 80))
        self._hdr_bot_lbl = tk.Label(hdr, text="Strip 2 — bottom half", font=("Segoe UI", 9, "bold"))
        self._hdr_bot_lbl.pack(side="left")

        img_row = tk.Frame(right)
        img_row.pack(fill="both", expand=True, pady=(6, 0))

        self._canvas_top = tk.Canvas(img_row, bg="white", relief="sunken",
                                      width=240, height=500)
        self._canvas_top.pack(side="left", padx=(0, 12))

        self._canvas_bot = tk.Canvas(img_row, bg="white", relief="sunken",
                                      width=240, height=500)
        self._canvas_bot.pack(side="left")

        # ── Page navigation (below preview) ────────────────────────────
        nav_row = tk.Frame(right)
        nav_row.pack(pady=(8, 0))
        self._prev_btn = tk.Button(
            nav_row, text="◀ Prev page", width=12,
            command=self._prev_page, state="disabled",
        )
        self._prev_btn.pack(side="left", padx=(0, 8))
        self._page_lbl = tk.Label(nav_row, text="Page 1 of 1", font=("Segoe UI", 9))
        self._page_lbl.pack(side="left", padx=(0, 8))
        self._next_btn = tk.Button(
            nav_row, text="Next page ▶", width=12,
            command=self._next_page, state="disabled",
        )
        self._next_btn.pack(side="left")

    # ── Courier list ───────────────────────────────────────────────────────

    def _refresh_courier_list(self) -> None:
        captured: dict = {}
        if CAPTURED_FILE.exists():
            try:
                captured = json.loads(CAPTURED_FILE.read_text(encoding="utf-8"))
            except Exception:
                pass

        self._courier_codes: list[str] = [
            code for code in COURIER_DISPLAY if captured.get(code)
        ]
        labels = [f"{COURIER_DISPLAY[c]}" for c in self._courier_codes]
        self._courier_combo["values"] = labels

        if labels:
            self._courier_combo.current(0)
            self._on_courier_changed()
        else:
            self._set_status(
                "No captured labels found in 'data/shipping/'.\n"
                "Book a shipment in the main app first — each courier's "
                "latest label is saved automatically."
            )

    # ── Event handlers ─────────────────────────────────────────────────────

    def _on_courier_changed(self, _event=None) -> None:
        idx = self._courier_combo.current()
        if idx < 0 or idx >= len(self._courier_codes):
            return
        self._courier_code = self._courier_codes[idx]
        self._page_num = 0
        self._strips_cache = None

        settings = _load_settings(self._courier_code)
        self._scale_var.set(round(settings["scale"] * 100))
        self._split_var.set(round(settings["split_ratio"] * 100))

        # no_split — update checkbox then call toggled handler to sync controls
        self._nosplit_var.set(settings.get("no_split", False))
        self._on_nosplit_toggled()

        # rotate_cw and skip_initial_rotate
        self._rotate_cw_var.set(settings.get("rotate_cw", False))
        self._skip_initial_rotate_var.set(settings.get("skip_initial_rotate", False))

        # label_length_mm
        lmm = settings.get("label_length_mm", 0.0)
        if lmm > 0:
            self._length_var.set(lmm)
            self._length_auto_var.set(False)
        else:
            self._length_auto_var.set(True)
        auto = self._length_auto_var.get()
        self._length_scale.configure(state="disabled" if auto else "normal")
        self._length_spin.configure(state="disabled" if auto else "normal")

        pdf_path = LABELS_DIR / f"{self._courier_code}.pdf"
        if pdf_path.exists():
            self._pdf_bytes = pdf_path.read_bytes()
            self._set_status(f"Loaded {pdf_path.name}  ({len(self._pdf_bytes):,} bytes)")
            self._schedule_preview()
        else:
            self._pdf_bytes = None
            self._set_status(f"PDF not found: {pdf_path}")

    def _on_param_changed(self, _val=None) -> None:
        self._strips_cache = None
        self._schedule_preview()

    def _on_nosplit_toggled(self) -> None:
        no_split = self._nosplit_var.get()
        state = "disabled" if no_split else "normal"
        self._split_scale.configure(state=state)
        self._split_spin.configure(state=state)
        if no_split:
            self._hdr_top_lbl.configure(text="Full label strip")
            self._hdr_bot_lbl.pack_forget()
            self._canvas_bot.pack_forget()
        else:
            self._hdr_top_lbl.configure(text="Strip 1 — top half")
            self._hdr_bot_lbl.pack(side="left")
            self._canvas_bot.pack(side="left")
        self._strips_cache = None
        self._schedule_preview()

    def _on_length_auto_toggled(self) -> None:
        auto = self._length_auto_var.get()
        state = "disabled" if auto else "normal"
        self._length_scale.configure(state=state)
        self._length_spin.configure(state=state)
        self._strips_cache = None
        self._schedule_preview()

    def _schedule_preview(self) -> None:
        """Debounce rapid slider updates — only redraw after 150 ms of quiet."""
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(150, self._update_preview)

    # ── Page navigation ────────────────────────────────────────────────────

    def _prev_page(self) -> None:
        if self._page_num > 0:
            self._page_num -= 1
            self._update_preview(use_cache=True)

    def _next_page(self) -> None:
        if self._page_num < self._page_count - 1:
            self._page_num += 1
            self._update_preview(use_cache=True)

    def _update_page_nav(self) -> None:
        self._page_lbl.configure(text=f"Page {self._page_num + 1} of {self._page_count}")
        self._prev_btn.configure(state="normal" if self._page_num > 0 else "disabled")
        self._next_btn.configure(
            state="normal" if self._page_num < self._page_count - 1 else "disabled"
        )

    # ── Preview ────────────────────────────────────────────────────────────

    def _update_preview(self, use_cache: bool = False) -> None:
        self._after_id = None
        if not self._pdf_bytes:
            return
        try:
            from src.shipping.label_printer import process_label_pdf
            from PIL import ImageTk

            scale = self._scale_var.get() / 100.0
            split_ratio = self._split_var.get() / 100.0
            no_split = self._nosplit_var.get()
            rotate_cw = self._rotate_cw_var.get()
            skip_initial_rotate = self._skip_initial_rotate_var.get()
            label_length_mm = 0.0 if self._length_auto_var.get() else self._length_var.get()

            if not use_cache or self._strips_cache is None:
                self._strips_cache = process_label_pdf(
                    self._pdf_bytes, scale=scale, split_ratio=split_ratio, no_split=no_split,
                    label_length_mm=label_length_mm, rotate_cw=rotate_cw,
                    skip_initial_rotate=skip_initial_rotate,
                )

            strips = self._strips_cache
            if not strips:
                return

            self._page_count = len(strips)
            self._page_num = min(self._page_num, self._page_count - 1)
            self._update_page_nav()

            top_img, bot_img = strips[self._page_num]

            # Fit images to fixed bounds — do NOT use winfo_width/height because
            # the canvas is resized to the result each frame, creating a feedback
            # loop that permanently shrinks the canvas when height decreases.
            fit_scale = min(240 / top_img.width, 500 / top_img.height)
            pw = max(1, int(top_img.width * fit_scale))
            ph = max(1, int(top_img.height * fit_scale))

            photo_top = ImageTk.PhotoImage(top_img.resize((pw, ph)))
            self._preview_photos = [photo_top]
            self._canvas_top.config(width=pw, height=ph)
            self._canvas_top.delete("all")
            self._canvas_top.create_image(0, 0, anchor="nw", image=photo_top)

            if bot_img is not None:
                photo_bot = ImageTk.PhotoImage(bot_img.resize((pw, ph)))
                self._preview_photos.append(photo_bot)
                self._canvas_bot.config(width=pw, height=ph)
                self._canvas_bot.delete("all")
                self._canvas_bot.create_image(0, 0, anchor="nw", image=photo_bot)

            length_info = f", length={label_length_mm:.0f}mm" if label_length_mm > 0 else ", length=auto"
            status = f"Strip size: {top_img.width}×{top_img.height} px  (scale={scale:.0%}"
            if no_split:
                status += f", no split{length_info})"
            else:
                flags = []
                if rotate_cw:
                    flags.append("rotate CW")
                if skip_initial_rotate:
                    flags.append("skip initial rotate")
                flag_str = ("  " + "  ".join(flags)) if flags else ""
                status += f", split={split_ratio:.0%}{length_info}{flag_str})"
            self._set_status(status)
        except Exception as exc:
            self._set_status(f"Preview error: {exc}")

    # ── Actions ────────────────────────────────────────────────────────────

    def _print_test(self) -> None:
        if not self._pdf_bytes:
            self._set_status("No label loaded — select a courier first.")
            return
        scale = self._scale_var.get() / 100.0
        split_ratio = self._split_var.get() / 100.0
        no_split = self._nosplit_var.get()
        rotate_cw = self._rotate_cw_var.get()
        skip_initial_rotate = self._skip_initial_rotate_var.get()
        label_length_mm = 0.0 if self._length_auto_var.get() else self._length_var.get()
        self._set_status("Sending to printer…")
        self.update_idletasks()

        def _do() -> None:
            try:
                from src.shipping.label_printer import print_label
                err = print_label(
                    self._pdf_bytes, scale=scale, split_ratio=split_ratio, no_split=no_split,
                    label_length_mm=label_length_mm, rotate_cw=rotate_cw,
                    skip_initial_rotate=skip_initial_rotate,
                )
                if err:
                    self.after(0, lambda: self._set_status(f"Print failed:\n{err}"))
                else:
                    self.after(0, lambda: self._set_status("Printed successfully."))
            except Exception as exc:
                self.after(0, lambda: self._set_status(f"Error: {exc}"))

        threading.Thread(target=_do, daemon=True).start()

    def _save_settings_ui(self) -> None:
        if not self._courier_code:
            self._set_status("No courier selected.")
            return
        scale = self._scale_var.get() / 100.0
        split_ratio = self._split_var.get() / 100.0
        no_split = self._nosplit_var.get()
        rotate_cw = self._rotate_cw_var.get()
        skip_initial_rotate = self._skip_initial_rotate_var.get()
        label_length_mm = 0.0 if self._length_auto_var.get() else self._length_var.get()
        try:
            _save_settings(self._courier_code, scale, split_ratio, label_length_mm,
                           no_split, rotate_cw, skip_initial_rotate)
            name = COURIER_DISPLAY.get(self._courier_code, self._courier_code)
            length_str = f"{label_length_mm:.0f}mm" if label_length_mm > 0 else "auto"
            flags = []
            if no_split:
                flags.append("no split")
            if rotate_cw:
                flags.append("rotate CW")
            if skip_initial_rotate:
                flags.append("skip initial rotate")
            flag_str = ("  " + "  ".join(flags)) if flags else ""
            self._set_status(
                f"Saved for {name}:\n  scale={scale:.0%}  split={split_ratio:.0%}"
                f"  length={length_str}{flag_str}\n"
                "These values will be used by the main app."
            )
        except Exception as exc:
            self._set_status(f"Save failed: {exc}")

    def _set_status(self, msg: str) -> None:
        self._status_lbl.configure(text=msg)


if __name__ == "__main__":
    # Change to project root so relative paths (data/shipping/, src/) work
    os.chdir(Path(__file__).resolve().parent.parent)
    app = LabelTuner()
    app.mainloop()
