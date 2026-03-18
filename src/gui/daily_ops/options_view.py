from __future__ import annotations

from datetime import datetime, timedelta

import customtkinter as ctk

# Mirrors the channel map from orders_tab.py
_NETO_CHANNEL_MAP: dict[str, list[str]] = {
    "Website":         ["Website"],
    "eBay (via Neto)": ["eBay"],
    "BigW":            ["BigW"],
    "Kogan":           ["Kogan"],
    "Amazon":          ["Amazon AU", "Amazon"],
}

_FILTER_KEYS = ["include_express", "include_click_collect", "ebay_direct"]

_FILTER_LABELS = {
    "include_express":       "Include Express orders",
    "include_click_collect": "Include Click & Collect orders",
    "ebay_direct":           "eBay (direct)",
}

_FILTER_DEFAULTS = {
    "include_express":       True,
    "include_click_collect": False,
    "ebay_direct":           True,
}


class OptionsView(ctk.CTkFrame):
    """
    Step 1 — Options screen for Daily Operations.

    User configures:
      - Date range (defaults to today only)
      - Platform toggles (Neto channels + eBay direct)
      - Filter toggles (express, click & collect)
      - Note filter phrases (add/remove/edit)

    Calls on_generate(options_dict) when Generate List is clicked.
    """

    def __init__(self, master, window, on_generate, on_back=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_generate = on_generate
        self._on_back = on_back
        self._platform_switches: dict[str, ctk.CTkSwitch] = {}
        self._filter_switches: dict[str, ctk.CTkSwitch] = {}
        self._phrase_vars: list[ctk.StringVar] = []
        self._phrase_rows: list[ctk.CTkFrame] = []
        self._build_ui()
        self._load_saved_toggles()

    # ── Build ────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 16, "pady": 6}

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # ── Date range ────────────────────────────────────────────────
        date_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        date_frame.pack(fill="x", **pad)

        ctk.CTkLabel(date_frame, text="Date range:", font=ctk.CTkFont(size=13, weight="bold")).pack(
            side="left"
        )

        today = datetime.today()
        self._from_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="DD/MM/YYYY")
        self._from_entry.insert(0, (today - timedelta(days=3)).strftime("%d/%m/%Y"))
        self._from_entry.pack(side="left", padx=(10, 6))

        ctk.CTkLabel(date_frame, text="to", font=ctk.CTkFont(size=13)).pack(side="left")

        self._to_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="DD/MM/YYYY")
        self._to_entry.insert(0, today.strftime("%d/%m/%Y"))
        self._to_entry.pack(side="left", padx=(6, 0))

        # ── Platform toggles ──────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="Platforms:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w"
        ).pack(fill="x", padx=16, pady=(12, 2))

        plat_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        plat_frame.pack(fill="x", padx=16, pady=(0, 4))

        for label in _NETO_CHANNEL_MAP:
            sw = ctk.CTkSwitch(
                plat_frame,
                text=label,
                font=ctk.CTkFont(size=12),
                width=50,
                command=self._save_toggles,
            )
            sw.select()
            sw.pack(side="left", padx=(0, 16))
            self._platform_switches[label] = sw

        # ── Filter toggles ────────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="Filters:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w"
        ).pack(fill="x", padx=16, pady=(12, 2))

        filter_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        filter_frame.pack(fill="x", padx=16, pady=(0, 4))

        for key in _FILTER_KEYS:
            sw = ctk.CTkSwitch(
                filter_frame,
                text=_FILTER_LABELS[key],
                font=ctk.CTkFont(size=12),
                width=50,
                command=self._save_toggles,
            )
            if _FILTER_DEFAULTS[key]:
                sw.select()
            else:
                sw.deselect()
            sw.pack(side="left", padx=(0, 16))
            self._filter_switches[key] = sw

        # ── Note filter phrases ───────────────────────────────────────
        phrase_header = ctk.CTkFrame(scroll, fg_color="transparent")
        phrase_header.pack(fill="x", padx=16, pady=(12, 2))

        ctk.CTkLabel(
            phrase_header,
            text="Note filter phrases  (orders containing these will be excluded):",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left")

        ctk.CTkButton(
            phrase_header,
            text="+ Add",
            width=60,
            height=24,
            font=ctk.CTkFont(size=11),
            command=self._add_phrase_row,
        ).pack(side="right")

        self._phrase_container = ctk.CTkFrame(scroll, fg_color="transparent")
        self._phrase_container.pack(fill="x", padx=16, pady=(0, 4))

        # Load existing phrases from config
        for phrase in self._window.config.app.note_filter_phrases:
            self._add_phrase_row(phrase)

        # ── Action row ────────────────────────────────────────────────
        action = ctk.CTkFrame(scroll, fg_color="transparent")
        action.pack(fill="x", padx=16, pady=(20, 12))

        self._error_label = ctk.CTkLabel(
            action,
            text="",
            text_color="red",
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self._error_label.pack(fill="x", pady=(0, 8))

        btn_row = ctk.CTkFrame(action, fg_color="transparent")
        btn_row.pack(fill="x")

        if self._on_back:
            ctk.CTkButton(
                btn_row,
                text="← Back to Menu",
                width=150,
                height=44,
                fg_color=("gray70", "gray30"),
                hover_color=("gray60", "gray25"),
                command=self._on_back,
            ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            btn_row,
            text="Generate List  →",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            command=self._on_generate_clicked,
        ).pack(side="left", fill="x", expand=True)

    # ── Phrase rows ───────────────────────────────────────────────────────

    def _add_phrase_row(self, initial_text: str = ""):
        row = ctk.CTkFrame(self._phrase_container, fg_color="transparent")
        row.pack(fill="x", pady=2)

        var = ctk.StringVar(value=initial_text)
        entry = ctk.CTkEntry(row, textvariable=var, width=300, font=ctk.CTkFont(size=12))
        entry.pack(side="left")

        def _remove(r=row, v=var):
            r.destroy()
            if v in self._phrase_vars:
                self._phrase_vars.remove(v)
            if r in self._phrase_rows:
                self._phrase_rows.remove(r)
            self._save_phrases()

        ctk.CTkButton(
            row, text="✕", width=28, height=28,
            font=ctk.CTkFont(size=11),
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=_remove,
        ).pack(side="left", padx=(6, 0))

        self._phrase_vars.append(var)
        self._phrase_rows.append(row)

        # Save whenever text changes
        var.trace_add("write", lambda *_: self._save_phrases())

    # ── Persistence ───────────────────────────────────────────────────────

    def _load_saved_toggles(self):
        saved = self._window.config.app.daily_ops_toggles
        if not saved:
            return
        for label, sw in self._platform_switches.items():
            if saved.get("platforms", {}).get(label, True):
                sw.select()
            else:
                sw.deselect()
        for key, sw in self._filter_switches.items():
            val = saved.get(key, _FILTER_DEFAULTS[key])
            if val:
                sw.select()
            else:
                sw.deselect()

    def _save_toggles(self):
        platforms = {label: (sw.get() == 1) for label, sw in self._platform_switches.items()}
        filters = {key: (sw.get() == 1) for key, sw in self._filter_switches.items()}
        toggles = {"platforms": platforms, **filters}
        self._window.config.app.daily_ops_toggles = toggles
        self._window.config._raw.setdefault("app", {})["daily_ops_toggles"] = toggles
        self._window.config.save()

    def _save_phrases(self):
        phrases = [v.get().strip() for v in self._phrase_vars if v.get().strip()]
        self._window.config.app.note_filter_phrases = phrases
        self._window.config._raw.setdefault("app", {})["note_filter_phrases"] = phrases
        self._window.config.save()

    # ── Generate ──────────────────────────────────────────────────────────

    def _on_generate_clicked(self):
        self._error_label.configure(text="")

        date_from, date_to = self._parse_dates()
        if date_from is None:
            return

        self._save_toggles()
        self._save_phrases()

        options = {
            "date_from": date_from,
            "date_to": date_to,
            "platforms": {
                label: (sw.get() == 1) for label, sw in self._platform_switches.items()
            },
            "ebay_direct": self._filter_switches["ebay_direct"].get() == 1,
            "include_express": self._filter_switches["include_express"].get() == 1,
            "include_click_collect": self._filter_switches["include_click_collect"].get() == 1,
            "note_filter_phrases": [
                v.get().strip() for v in self._phrase_vars if v.get().strip()
            ],
        }
        self._on_generate(options)

    def _parse_dates(self):
        from_str = self._from_entry.get().strip()
        to_str = self._to_entry.get().strip()
        try:
            date_from = datetime.strptime(from_str, "%d/%m/%Y")
        except ValueError:
            self._error_label.configure(text=f"Invalid 'From' date: '{from_str}'. Use DD/MM/YYYY.")
            return None, None
        try:
            date_to = datetime.strptime(to_str, "%d/%m/%Y")
        except ValueError:
            self._error_label.configure(text=f"Invalid 'To' date: '{to_str}'. Use DD/MM/YYYY.")
            return None, None
        if date_from > date_to:
            self._error_label.configure(text="'From' date must be on or before 'To' date.")
            return None, None
        return date_from, date_to
