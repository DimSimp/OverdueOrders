from __future__ import annotations

from datetime import datetime, timedelta

import customtkinter as ctk

# Mirrors the channel map from options_view.py
_NETO_CHANNEL_MAP: dict[str, list[str]] = {
    "Website":          ["Website"],
    "eBay (via Neto)":  ["eBay"],
    "BigW":             ["BigW"],
    "Kogan":            ["Kogan"],
    "Amazon":           ["Amazon AU", "Amazon"],
    "Everydaymarket":   ["Everydaymarket"],
    "Control Panel":    ["Control Panel"],
    "Quote":            ["Quote"],
}


class AllOrdersOptionsView(ctk.CTkFrame):
    """
    Step 1 — Options screen for Show All Orders.

    User configures:
      - Date range (defaults to today -3 days to today)
      - Platform toggles (Neto channels + eBay direct)

    No express / click-collect / note-phrase filters — we fetch everything.

    Calls on_fetch(options_dict) when Fetch Orders is clicked.
    """

    def __init__(self, master, window, on_fetch, on_back=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_fetch = on_fetch
        self._on_back = on_back
        self._platform_switches: dict[str, ctk.CTkSwitch] = {}
        self._ebay_direct_switch: ctk.CTkSwitch | None = None
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

        # ── eBay Direct toggle ────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="Additional sources:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w"
        ).pack(fill="x", padx=16, pady=(12, 2))

        extra_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        extra_frame.pack(fill="x", padx=16, pady=(0, 4))

        self._ebay_direct_switch = ctk.CTkSwitch(
            extra_frame,
            text="eBay (direct)",
            font=ctk.CTkFont(size=12),
            width=50,
            command=self._save_toggles,
        )
        self._ebay_direct_switch.select()
        self._ebay_direct_switch.pack(side="left", padx=(0, 16))

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
            text="Fetch Orders  →",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            command=self._on_fetch_clicked,
        ).pack(side="left", fill="x", expand=True)

    # ── Persistence ───────────────────────────────────────────────────────

    def _load_saved_toggles(self):
        saved = self._window.config.app.all_orders_toggles
        if not saved:
            return
        for label, sw in self._platform_switches.items():
            if saved.get("platforms", {}).get(label, True):
                sw.select()
            else:
                sw.deselect()
        ebay_direct = saved.get("ebay_direct", True)
        if self._ebay_direct_switch:
            if ebay_direct:
                self._ebay_direct_switch.select()
            else:
                self._ebay_direct_switch.deselect()

    def _save_toggles(self):
        platforms = {label: (sw.get() == 1) for label, sw in self._platform_switches.items()}
        ebay_direct = (self._ebay_direct_switch.get() == 1) if self._ebay_direct_switch else True
        toggles = {"platforms": platforms, "ebay_direct": ebay_direct}
        self._window.config.app.all_orders_toggles = toggles
        self._window.config._raw.setdefault("app", {})["all_orders_toggles"] = toggles
        self._window.config.save()

    # ── Fetch ──────────────────────────────────────────────────────────────

    def _on_fetch_clicked(self):
        self._error_label.configure(text="")

        date_from, date_to = self._parse_dates()
        if date_from is None:
            return

        self._save_toggles()

        options = {
            "date_from": date_from,
            "date_to": date_to,
            "platforms": {
                label: (sw.get() == 1) for label, sw in self._platform_switches.items()
            },
            "ebay_direct": (self._ebay_direct_switch.get() == 1) if self._ebay_direct_switch else True,
        }
        self._on_fetch(options)

    def _parse_dates(self):
        from_str = self._from_entry.get().strip()
        to_str = self._to_entry.get().strip()
        try:
            date_from = datetime.strptime(from_str, "%d/%m/%Y")
        except ValueError:
            self._error_label.configure(text=f"Invalid 'From' date: '{from_str}'. Use DD/MM/YYYY.")
            return None, None
        try:
            _parsed_to = datetime.strptime(to_str, "%d/%m/%Y")
            today = datetime.now().date()
            if _parsed_to.date() >= today:
                date_to = datetime.now()
            else:
                date_to = _parsed_to.replace(hour=23, minute=59, second=59)
        except ValueError:
            self._error_label.configure(text=f"Invalid 'To' date: '{to_str}'. Use DD/MM/YYYY.")
            return None, None
        if date_from > date_to:
            self._error_label.configure(text="'From' date must be on or before 'To' date.")
            return None, None
        return date_from, date_to
