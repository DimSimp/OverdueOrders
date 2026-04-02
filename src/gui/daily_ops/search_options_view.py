from __future__ import annotations

import logging
import re
import threading
from datetime import datetime, timedelta

import customtkinter as ctk

from src.ebay_client import EbayAuthError, EbayAPIError
from src.neto_client import NetoAPIError

log = logging.getLogger(__name__)

# ── Constants ─────────────────────────────────────────────────────────────────

# Order number format patterns for fast-path routing and validation.
# Neto: "N" followed by exactly 6 digits (e.g. N100123)
# eBay: two digits, hyphen, five digits, hyphen, five digits (e.g. 12-12345-67890)
_NETO_ORDER_RE = re.compile(r"^N\d{6}$")
_EBAY_ORDER_RE = re.compile(r"^\d{2}-\d{5}-\d{5}$")

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

_NETO_STATUSES: list[str] = [
    "New", "Pick", "Pack", "Pending Dispatch", "Pending Pickup",
    "Dispatched", "On Hold", "Uncommitted", "New Backorder",
    "Backorder Approved", "Cancelled",
]

# Maps eBay status label → list of API fulfillment status values
_EBAY_STATUSES: dict[str, list[str]] = {
    "Awaiting Dispatch": ["NOT_STARTED", "IN_PROGRESS"],
    "Dispatched":        ["FULFILLED"],
}

# All Neto statuses default ON; both eBay statuses default ON
_NETO_STATUS_DEFAULTS: dict[str, bool] = {s: True for s in _NETO_STATUSES}
_EBAY_STATUS_DEFAULTS: dict[str, bool] = {lbl: True for lbl in _EBAY_STATUSES}


# ── Shared filter helpers (also imported by search_results_view) ───────────────

def _filter_neto_by_channel(orders: list, options: dict) -> list:
    """Client-side channel filter — mirrors AllOrdersFetchView._filter_neto_by_channel."""
    platforms = options.get("platforms", {})
    channel_enabled: dict[str, bool] = {}
    for label, channels in _NETO_CHANNEL_MAP.items():
        is_on = platforms.get(label, True)
        for ch in channels:
            channel_enabled[ch.lower()] = is_on
    result = []
    for order in orders:
        ch = (order.sales_channel or "").lower()
        if ch in channel_enabled:
            if channel_enabled[ch]:
                result.append(order)
        else:
            result.append(order)
    return result


def _apply_text_filters(
    neto_orders: list,
    ebay_orders: list,
    terms: dict,
) -> tuple[list, list]:
    """
    Apply partial case-insensitive substring matches across all 4 search fields.
    Empty terms match everything.
    """
    order_number = terms.get("order_number", "").lower()
    sku = terms.get("sku", "").lower()
    title = terms.get("title", "").lower()
    name = terms.get("name", "").lower()

    def _neto_matches(o) -> bool:
        if order_number:
            id_lower = o.order_id.lower()
            po_lower = (o.purchase_order_number or "").lower()
            if order_number not in id_lower and order_number not in po_lower:
                return False
        if name:
            customer = (
                f"{o.ship_first_name} {o.ship_last_name}".strip()
                or o.customer_name or ""
            ).lower()
            if name not in customer:
                return False
        if sku:
            if not any(sku in (li.sku or "").lower() for li in o.line_items):
                return False
        if title:
            if not any(title in (li.product_name or "").lower() for li in o.line_items):
                return False
        return True

    def _ebay_matches(o) -> bool:
        if order_number:
            if order_number not in o.order_id.lower():
                return False
        if name:
            customer = (o.ship_name or o.buyer_name or "").lower()
            if name not in customer:
                return False
        if sku:
            if not any(sku in (li.sku or "").lower() for li in o.line_items):
                return False
        if title:
            if not any(title in (li.title or "").lower() for li in o.line_items):
                return False
        return True

    return (
        [o for o in neto_orders if _neto_matches(o)],
        [o for o in ebay_orders if _ebay_matches(o)],
    )


# ── View ──────────────────────────────────────────────────────────────────────

class SearchOptionsView(ctk.CTkFrame):
    """
    Search for Order — options form with inline fetch progress.

    Form state: shows all search options and fields.
    Loading state: form dims/disables, progress bar appears.
    On 0 results: re-enables form, shows error message (no navigation).
    On results: stores orders on window, calls on_results().
    """

    def __init__(self, master, window, on_results, on_back=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_results = on_results
        self._on_back = on_back

        self._platform_switches: dict[str, ctk.CTkSwitch] = {}
        self._ebay_direct_switch: ctk.CTkSwitch | None = None
        self._neto_status_vars: dict[str, ctk.BooleanVar] = {}
        self._neto_status_cbs: list[ctk.CTkCheckBox] = []
        self._ebay_status_vars: dict[str, ctk.BooleanVar] = {}
        self._ebay_status_cbs: list[ctk.CTkCheckBox] = []

        # Widgets that get disabled during search (populated in _build_ui)
        self._form_entries: list = []

        self._build_ui()
        self._load_saved_toggles()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=0, pady=0)
        self._scroll = scroll

        pad = {"padx": 16, "pady": 6}

        # ── Date range ────────────────────────────────────────────────────
        date_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        date_frame.pack(fill="x", **pad)

        ctk.CTkLabel(
            date_frame, text="Date range:", font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        today = datetime.today()
        self._from_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="DD/MM/YYYY")
        self._from_entry.insert(0, (today - timedelta(days=90)).strftime("%d/%m/%Y"))
        self._from_entry.pack(side="left", padx=(10, 6))
        self._form_entries.append(self._from_entry)

        ctk.CTkLabel(date_frame, text="to", font=ctk.CTkFont(size=13)).pack(side="left")

        self._to_entry = ctk.CTkEntry(date_frame, width=110, placeholder_text="DD/MM/YYYY")
        self._to_entry.insert(0, today.strftime("%d/%m/%Y"))
        self._to_entry.pack(side="left", padx=(6, 0))
        self._form_entries.append(self._to_entry)

        # ── Platform toggles ──────────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="Platforms:", font=ctk.CTkFont(size=13, weight="bold"), anchor="w"
        ).pack(fill="x", padx=16, pady=(12, 2))

        plat_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        plat_frame.pack(fill="x", padx=16, pady=(0, 4))

        for label in _NETO_CHANNEL_MAP:
            sw = ctk.CTkSwitch(
                plat_frame, text=label, font=ctk.CTkFont(size=12),
                width=50, command=self._save_toggles,
            )
            sw.select()
            sw.pack(side="left", padx=(0, 16))
            self._platform_switches[label] = sw

        # eBay direct
        ctk.CTkLabel(
            scroll, text="Additional sources:",
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        ).pack(fill="x", padx=16, pady=(12, 2))

        extra_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        extra_frame.pack(fill="x", padx=16, pady=(0, 4))

        self._ebay_direct_switch = ctk.CTkSwitch(
            extra_frame, text="eBay (direct)", font=ctk.CTkFont(size=12),
            width=50, command=self._save_toggles,
        )
        self._ebay_direct_switch.select()
        self._ebay_direct_switch.pack(side="left", padx=(0, 16))

        # ── Neto order status ─────────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="Neto order status:",
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        ).pack(fill="x", padx=16, pady=(12, 2))

        neto_status_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        neto_status_frame.pack(fill="x", padx=16, pady=(0, 4))

        for i, status in enumerate(_NETO_STATUSES):
            var = ctk.BooleanVar(value=_NETO_STATUS_DEFAULTS.get(status, True))
            self._neto_status_vars[status] = var
            cb = ctk.CTkCheckBox(
                neto_status_frame, text=status, variable=var,
                font=ctk.CTkFont(size=12), width=10,
                command=self._save_toggles,
            )
            cb.grid(row=i // 4, column=i % 4, sticky="w", padx=(0, 16), pady=2)
            self._neto_status_cbs.append(cb)

        # ── eBay order status ─────────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="eBay order status:",
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        ).pack(fill="x", padx=16, pady=(12, 2))

        ebay_status_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        ebay_status_frame.pack(fill="x", padx=16, pady=(0, 4))

        for label in _EBAY_STATUSES:
            var = ctk.BooleanVar(value=_EBAY_STATUS_DEFAULTS.get(label, True))
            self._ebay_status_vars[label] = var
            cb = ctk.CTkCheckBox(
                ebay_status_frame, text=label, variable=var,
                font=ctk.CTkFont(size=12), width=10,
                command=self._save_toggles,
            )
            cb.pack(side="left", padx=(0, 16))
            self._ebay_status_cbs.append(cb)

        # ── Search fields ──────────────────────────────────────────────────
        ctk.CTkLabel(
            scroll, text="Search Terms:",
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        ).pack(fill="x", padx=16, pady=(16, 4))

        fields_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        fields_frame.pack(fill="x", padx=16, pady=(0, 4))

        field_defs = [
            ("Order Number:", "_order_no_entry"),
            ("Item SKU:",     "_sku_entry"),
            ("Item Title:",   "_title_entry"),
            ("Name:",         "_name_entry"),
        ]

        for row_idx, (label_text, attr_name) in enumerate(field_defs):
            ctk.CTkLabel(
                fields_frame, text=label_text,
                font=ctk.CTkFont(size=12), anchor="e", width=100,
            ).grid(row=row_idx, column=0, sticky="e", padx=(0, 8), pady=4)
            entry = ctk.CTkEntry(fields_frame, width=320)
            entry.grid(row=row_idx, column=1, sticky="w", pady=4)
            setattr(self, attr_name, entry)
            self._form_entries.append(entry)

        # ── Progress / error ───────────────────────────────────────────────
        progress_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        progress_frame.pack(fill="x", padx=16, pady=(12, 0))

        self._progress_bar = ctk.CTkProgressBar(progress_frame, mode="indeterminate")
        # Hidden initially — shown when search starts
        self._progress_bar_visible = False

        self._progress_label = ctk.CTkLabel(
            progress_frame, text="", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"), anchor="w",
        )
        self._progress_label.pack(fill="x")

        self._error_label = ctk.CTkLabel(
            scroll, text="", font=ctk.CTkFont(size=12),
            text_color="red", anchor="w",
        )
        self._error_label.pack(fill="x", padx=16, pady=(4, 0))

        # ── Action buttons ─────────────────────────────────────────────────
        action_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        action_frame.pack(fill="x", padx=16, pady=(16, 12))

        if self._on_back:
            self._back_btn = ctk.CTkButton(
                action_frame, text="← Back to Menu", width=150, height=44,
                fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
                command=self._on_back,
            )
            self._back_btn.pack(side="left", padx=(0, 12))

        self._search_btn = ctk.CTkButton(
            action_frame, text="Search  ▶",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44, command=self._on_search_clicked,
        )
        self._search_btn.pack(side="left", fill="x", expand=True)

    # ── Persistence ───────────────────────────────────────────────────────────

    def _load_saved_toggles(self):
        saved = self._window.config.app.search_order_toggles
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
        for status, var in self._neto_status_vars.items():
            default = _NETO_STATUS_DEFAULTS.get(status, True)
            var.set(saved.get("neto_statuses", {}).get(status, default))
        for label, var in self._ebay_status_vars.items():
            default = _EBAY_STATUS_DEFAULTS.get(label, True)
            var.set(saved.get("ebay_statuses", {}).get(label, default))

    def _save_toggles(self):
        platforms = {label: (sw.get() == 1) for label, sw in self._platform_switches.items()}
        ebay_direct = (self._ebay_direct_switch.get() == 1) if self._ebay_direct_switch else True
        neto_statuses = {s: var.get() for s, var in self._neto_status_vars.items()}
        ebay_statuses = {lbl: var.get() for lbl, var in self._ebay_status_vars.items()}
        toggles = {
            "platforms": platforms,
            "ebay_direct": ebay_direct,
            "neto_statuses": neto_statuses,
            "ebay_statuses": ebay_statuses,
        }
        self._window.config.app.search_order_toggles = toggles
        self._window.config._raw.setdefault("app", {})["search_order_toggles"] = toggles
        self._window.config.save()

    # ── Search ────────────────────────────────────────────────────────────────

    def _on_search_clicked(self):
        self._error_label.configure(text="")

        date_from, date_to = self._parse_dates()
        if date_from is None:
            return

        order_no = self._order_no_entry.get().strip()
        sku = self._sku_entry.get().strip()
        title = self._title_entry.get().strip()
        name = self._name_entry.get().strip()

        if not any([order_no, sku, title, name]):
            self._error_label.configure(text="Please enter at least one search term.")
            return

        # Validate order number format if provided, and determine fast-path platform.
        # A recognised exact ID bypasses the full date-range scan entirely.
        order_id_fast_path = None
        if order_no:
            if _NETO_ORDER_RE.match(order_no):
                order_id_fast_path = {"platform": "neto", "order_id": order_no}
            elif _EBAY_ORDER_RE.match(order_no):
                order_id_fast_path = {"platform": "ebay", "order_id": order_no}
            else:
                self._error_label.configure(
                    text=(
                        "Unrecognised order number format. "
                        "Neto orders look like N123456, eBay orders like 12-12345-67890."
                    )
                )
                return

        # Collect selected statuses / platforms (used for normal-path searches)
        neto_statuses = [s for s, var in self._neto_status_vars.items() if var.get()]
        ebay_statuses_on = [lbl for lbl, var in self._ebay_status_vars.items() if var.get()]
        ebay_fulfillment_statuses: list[str] = []
        for lbl in ebay_statuses_on:
            ebay_fulfillment_statuses.extend(_EBAY_STATUSES.get(lbl, []))

        platforms = {label: (sw.get() == 1) for label, sw in self._platform_switches.items()}
        ebay_direct = (self._ebay_direct_switch.get() == 1) if self._ebay_direct_switch else True

        options = {
            "date_from": date_from,
            "date_to": date_to,
            "platforms": platforms,
            "ebay_direct": ebay_direct,
            "neto_statuses": neto_statuses,
            "ebay_fulfillment_statuses": ebay_fulfillment_statuses,
            "order_id_fast_path": order_id_fast_path,
            "search_terms": {
                "order_number": order_no,
                "sku": sku,
                "title": title,
                "name": name,
            },
        }

        self._save_toggles()
        self._start_search(options)

    def _start_search(self, options: dict):
        self._search_btn.configure(state="disabled")
        if hasattr(self, "_back_btn"):
            self._back_btn.configure(state="disabled")
        self._set_form_enabled(False)

        if not self._progress_bar_visible:
            self._progress_bar.pack(fill="x", pady=(0, 4), before=self._progress_label)
            self._progress_bar_visible = True
        self._progress_bar.start()
        self._progress_label.configure(
            text="Searching…", text_color=("gray50", "gray60")
        )

        threading.Thread(
            target=self._search_worker, args=(options,), daemon=True
        ).start()

    def _set_form_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for entry in self._form_entries:
            entry.configure(state=state)
        for sw in self._platform_switches.values():
            sw.configure(state=state)
        if self._ebay_direct_switch:
            self._ebay_direct_switch.configure(state=state)
        for cb in self._neto_status_cbs:
            cb.configure(state=state)
        for cb in self._ebay_status_cbs:
            cb.configure(state=state)

    def _search_worker(self, options: dict):
        fast_path = options.get("order_id_fast_path")
        if fast_path:
            self._fast_path_worker(options, fast_path)
            return

        # ── Normal path: run Neto + eBay in parallel ───────────────────────
        neto_orders: list = []
        ebay_orders: list = []
        neto_error: list[str] = []   # list used for mutation from inner functions
        ebay_warning: list[str] = []

        def _fetch_neto():
            neto_statuses = options.get("neto_statuses", [])
            any_neto_platform = any(options.get("platforms", {}).values())
            if not (neto_statuses and any_neto_platform):
                return
            try:
                raw = self._window.neto_client.search_orders(
                    options["date_from"], options["date_to"],
                    statuses=neto_statuses,
                )
                neto_orders.extend(_filter_neto_by_channel(raw, options))
                log.debug("Search: Neto returned %d → %d after channel filter",
                          len(raw), len(neto_orders))
            except NetoAPIError as exc:
                neto_error.append(f"Neto error: {exc}")
            except Exception as exc:
                neto_error.append(f"Neto fetch failed: {exc}")

        def _fetch_ebay():
            ebay_statuses = options.get("ebay_fulfillment_statuses", [])
            if not (options.get("ebay_direct", True) and ebay_statuses):
                return
            try:
                results = self._window.ebay_client.search_orders(
                    options["date_from"], options["date_to"],
                    fulfillment_statuses=ebay_statuses,
                    enrich_notes=False,
                )
                ebay_orders.extend(results)
                log.debug("Search: eBay returned %d orders", len(results))
            except EbayAuthError as exc:
                ebay_warning.append(f"eBay auth error: {exc}")
                log.warning("Search eBay auth error (continuing with Neto): %s", exc)
            except EbayAPIError as exc:
                ebay_warning.append(f"eBay API error: {exc}")
                log.warning("Search eBay API error (continuing with Neto): %s", exc)
            except Exception as exc:
                ebay_warning.append(f"eBay fetch failed: {exc}")
                log.warning("Search eBay fetch failed (continuing with Neto): %s", exc)

        neto_thread = threading.Thread(target=_fetch_neto, daemon=True)
        ebay_thread = threading.Thread(target=_fetch_ebay, daemon=True)
        neto_thread.start()
        ebay_thread.start()
        neto_thread.join()
        ebay_thread.join()

        if neto_error:
            self.after(0, lambda msg=neto_error[0]: self._on_search_error(msg))
            return

        terms = options.get("search_terms", {})
        filtered_neto, filtered_ebay = _apply_text_filters(neto_orders, ebay_orders, terms)
        warning = ebay_warning[0] if ebay_warning else ""
        self.after(
            0,
            lambda n=filtered_neto, e=filtered_ebay, w=warning, o=options:
                self._on_search_done(n, e, w, o),
        )

    def _fast_path_worker(self, options: dict, fast_path: dict):
        """Targeted single-order lookup when the user entered a fully-qualified order ID."""
        platform = fast_path["platform"]
        order_id = fast_path["order_id"]
        neto_orders: list = []
        ebay_orders: list = []
        error_msg: str | None = None
        ebay_warning: str = ""

        self.after(
            0,
            lambda oid=order_id: self._progress_label.configure(
                text=f"Looking up order {oid}…"
            ),
        )

        if platform == "neto":
            try:
                neto_orders = self._window.neto_client.get_order_by_exact_id(order_id)
            except NetoAPIError as exc:
                error_msg = f"Neto error: {exc}"
            except Exception as exc:
                error_msg = f"Neto fetch failed: {exc}"
        else:  # ebay
            try:
                ebay_orders = self._window.ebay_client.get_order_by_exact_id(order_id)
            except EbayAuthError as exc:
                ebay_warning = f"eBay auth: {exc}"
            except EbayAPIError as exc:
                error_msg = f"eBay error: {exc}"
            except Exception as exc:
                error_msg = f"eBay fetch failed: {exc}"

        if error_msg:
            self.after(0, lambda msg=error_msg: self._on_search_error(msg))
            return

        # Apply remaining filters (sku, title, name) — order_number already matched by ID
        terms = {k: v for k, v in options.get("search_terms", {}).items() if k != "order_number"}
        neto_orders, ebay_orders = _apply_text_filters(neto_orders, ebay_orders, terms)

        self.after(
            0,
            lambda n=neto_orders, e=ebay_orders, w=ebay_warning, o=options:
                self._on_search_done(n, e, w, o),
        )

    def _on_search_done(
        self,
        neto_orders: list,
        ebay_orders: list,
        ebay_warning: str,
        options: dict,
    ):
        self._progress_bar.stop()
        total = len(neto_orders) + len(ebay_orders)

        if total == 0:
            self._reset_to_form_state()
            msg = "No orders found matching the search criteria."
            if ebay_warning:
                msg += f"  (eBay: {ebay_warning})"
            self._error_label.configure(text=msg)
            return

        # Store results on window
        self._window.search_neto_orders = neto_orders
        self._window.search_ebay_orders = ebay_orders
        self._window.search_options = options

        self._reset_to_form_state()
        self._on_results()

    def _on_search_error(self, message: str):
        self._progress_bar.stop()
        self._reset_to_form_state()
        self._error_label.configure(text=message)

    def _reset_to_form_state(self):
        self._search_btn.configure(state="normal")
        if hasattr(self, "_back_btn"):
            self._back_btn.configure(state="normal")
        self._set_form_enabled(True)
        self._progress_label.configure(text="")
        if self._progress_bar_visible:
            self._progress_bar.pack_forget()
            self._progress_bar_visible = False

    # ── Date parsing ──────────────────────────────────────────────────────────

    def _parse_dates(self):
        from_str = self._from_entry.get().strip()
        to_str = self._to_entry.get().strip()
        try:
            date_from = datetime.strptime(from_str, "%d/%m/%Y")
        except ValueError:
            self._error_label.configure(
                text=f"Invalid 'From' date: '{from_str}'. Use DD/MM/YYYY."
            )
            return None, None
        try:
            _parsed_to = datetime.strptime(to_str, "%d/%m/%Y")
            today = datetime.now().date()
            if _parsed_to.date() >= today:
                date_to = datetime.now()
            else:
                date_to = _parsed_to.replace(hour=23, minute=59, second=59)
        except ValueError:
            self._error_label.configure(
                text=f"Invalid 'To' date: '{to_str}'. Use DD/MM/YYYY."
            )
            return None, None
        if date_from > date_to:
            self._error_label.configure(
                text="'From' date must be on or before 'To' date."
            )
            return None, None
        return date_from, date_to
