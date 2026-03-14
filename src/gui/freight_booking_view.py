from __future__ import annotations

import logging
import threading
import tkinter as tk

log = logging.getLogger("freight_booking")

import customtkinter as ctk

from src.neto_client import NetoClient
from src.shipping.models import (
    PACKAGE_PRESETS,
    Address,
    Package,
    Quote,
    ShipmentRequest,
    address_from_ebay_order,
    address_from_neto_order,
    sender_from_config,
)
from src.shipping.quote_engine import QuoteEngine


# Couriers that support express shipping
EXPRESS_CAPABLE_COURIERS = {"auspost", "allied", "bonds"}

# ── Courier registry ─────────────────────────────────────────────────────────

def _build_couriers(courier_configs: dict):
    """Instantiate courier objects from the shipping.couriers config dict."""
    from src.shipping.couriers.allied import AlliedCourier
    from src.shipping.couriers.aramex import AramexCourier
    from src.shipping.couriers.auspost import AusPostCourier
    from src.shipping.couriers.bonds import BondsCourier
    from src.shipping.couriers.dai_post import DaiPostCourier

    registry = {
        "auspost": AusPostCourier,
        "aramex": AramexCourier,
        "bonds": BondsCourier,
        "allied": AlliedCourier,
        "dai_post": DaiPostCourier,
    }
    couriers = []
    for code, cls in registry.items():
        cfg = courier_configs.get(code, {})
        if cfg.get("enabled", False):
            couriers.append(cls(cfg))
    return couriers


# ── Package Row Widget ────────────────────────────────────────────────────────

class PackageRow(ctk.CTkFrame):
    """A single row representing one package with dimensions + preset selector."""

    def __init__(self, master, sku: str = "", on_remove=None, **kwargs):
        super().__init__(master, fg_color=("gray92", "gray18"), corner_radius=6, **kwargs)
        self._sku = sku
        self._on_remove = on_remove
        self._auto_label: ctk.CTkLabel | None = None
        self._build()

    def _build(self):
        # Row 1: SKU label + preset dropdown + remove button
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=8, pady=(6, 2))

        if self._sku:
            ctk.CTkLabel(top, text=f"SKU: {self._sku}", font=ctk.CTkFont(size=11, weight="bold")).pack(
                side="left", padx=(0, 12)
            )

        ctk.CTkLabel(top, text="Preset:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 4))
        preset_names = list(PACKAGE_PRESETS.keys())
        self._preset_var = ctk.StringVar(value="Custom")
        self._preset_combo = ctk.CTkComboBox(
            top, values=preset_names, variable=self._preset_var,
            width=200, font=ctk.CTkFont(size=11),
            command=self._on_preset_change,
        )
        self._preset_combo.pack(side="left", padx=(0, 12))

        self._auto_label = ctk.CTkLabel(top, text="", font=ctk.CTkFont(size=10), text_color="gray50")
        self._auto_label.pack(side="left", padx=(0, 8))

        if self._on_remove:
            ctk.CTkButton(
                top, text="✕", width=28, height=28, fg_color="gray50",
                hover_color="red", command=self._on_remove,
            ).pack(side="right")

        # Row 2: Dimension entries
        dims = ctk.CTkFrame(self, fg_color="transparent")
        dims.pack(fill="x", padx=8, pady=(2, 6))

        self._weight = self._dim_field(dims, "Weight (kg)")
        self._length = self._dim_field(dims, "Length (cm)")
        self._width = self._dim_field(dims, "Width (cm)")
        self._height = self._dim_field(dims, "Height (cm)")

        # Cubic weight display
        self._cubic_label = ctk.CTkLabel(dims, text="", font=ctk.CTkFont(size=10), text_color="gray50")
        self._cubic_label.pack(side="left", padx=(12, 0))

        # Bind updates
        for entry in (self._weight, self._length, self._width, self._height):
            entry.bind("<KeyRelease>", lambda _: self._update_cubic())

    def _dim_field(self, parent, label: str) -> ctk.CTkEntry:
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=10)).pack(side="left", padx=(0, 2))
        entry = ctk.CTkEntry(parent, width=70, font=ctk.CTkFont(size=11))
        entry.pack(side="left", padx=(0, 10))
        return entry

    def _on_preset_change(self, preset_name: str):
        preset = PACKAGE_PRESETS.get(preset_name, {})
        if not preset:
            return
        self._set_values(
            weight=preset["weight_kg"],
            length=preset["length_cm"],
            width=preset["width_cm"],
            height=preset["height_cm"],
        )
        if self._auto_label:
            self._auto_label.configure(text="(preset)")

    def set_from_dimensions(self, dims: dict, source: str = "auto"):
        """Fill dimension fields from a dict with weight_kg/length_cm/width_cm/height_cm."""
        self._set_values(
            weight=dims.get("weight_kg", 0),
            length=dims.get("length_cm", 0),
            width=dims.get("width_cm", 0),
            height=dims.get("height_cm", 0),
        )
        self._preset_var.set("Custom")
        if self._auto_label:
            self._auto_label.configure(text=f"({source})")

    def _set_values(self, weight: float, length: float, width: float, height: float):
        for entry, val in [
            (self._weight, weight), (self._length, length),
            (self._width, width), (self._height, height),
        ]:
            entry.delete(0, "end")
            entry.insert(0, str(val))
        self._update_cubic()

    def _update_cubic(self):
        try:
            l = float(self._length.get() or 0)
            w = float(self._width.get() or 0)
            h = float(self._height.get() or 0)
            cubic = (l * w * h) / 4000
            self._cubic_label.configure(text=f"Cubic: {cubic:.2f} kg")
        except ValueError:
            self._cubic_label.configure(text="")

    def get_package(self) -> Package | None:
        """Parse entries into a Package, or None if invalid."""
        try:
            return Package(
                weight_kg=float(self._weight.get() or 0),
                length_cm=float(self._length.get() or 0),
                width_cm=float(self._width.get() or 0),
                height_cm=float(self._height.get() or 0),
            )
        except ValueError:
            return None


# ── Quote Card Widget ─────────────────────────────────────────────────────────

class QuoteCard(ctk.CTkFrame):
    """A clickable card showing a single courier quote."""

    def __init__(self, master, quote: Quote, on_select, is_cheapest: bool = False, **kwargs):
        border_color = ("green3", "green4") if is_cheapest else ("gray70", "gray35")
        fg = ("gray95", "gray16") if not quote.error else ("gray88", "gray22")
        super().__init__(master, fg_color=fg, border_width=2, border_color=border_color,
                         corner_radius=8, **kwargs)
        self._quote = quote
        self._on_select = on_select
        self._selected = False

        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=8)

        name_font = ctk.CTkFont(size=13, weight="bold")
        ctk.CTkLabel(row, text=quote.courier_name, font=name_font).pack(side="left", padx=(0, 8))

        if quote.error:
            ctk.CTkLabel(row, text=quote.error, font=ctk.CTkFont(size=11),
                         text_color="red", wraplength=400).pack(side="left", fill="x", expand=True)
        else:
            ctk.CTkLabel(row, text=quote.service_name, font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 12))
            price_text = f"${quote.price:.2f}"
            if is_cheapest:
                price_text += "  ★ Cheapest"
            ctk.CTkLabel(row, text=price_text, font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=("green4", "green3") if is_cheapest else None).pack(side="right", padx=(8, 0))
            if quote.estimated_days:
                ctk.CTkLabel(row, text=quote.estimated_days, font=ctk.CTkFont(size=10),
                             text_color="gray50").pack(side="right", padx=(0, 12))

        if not quote.error:
            self.bind("<Button-1>", lambda _: self._click())
            for child in self.winfo_children():
                child.bind("<Button-1>", lambda _: self._click())
                for grandchild in child.winfo_children():
                    grandchild.bind("<Button-1>", lambda _: self._click())

    def _click(self):
        self._on_select(self._quote)

    def set_selected(self, selected: bool):
        self._selected = selected
        if selected:
            self.configure(border_color=("dodgerblue", "dodgerblue3"))
        else:
            self.configure(border_color=("gray70", "gray35"))


# ── Main Freight Booking View ─────────────────────────────────────────────────

class FreightBookingView(ctk.CTkFrame):
    """
    Full-frame view for freight quoting. Replaces the OrderDetailView when
    the user clicks "Book Freight".
    """

    def __init__(
        self,
        master,
        *,
        order_id: str,
        platform: str,
        neto_order=None,
        ebay_order=None,
        neto_client: NetoClient | None = None,
        shipping_config,
        dry_run: bool = True,
        on_back,
        on_courier_selected,
    ):
        super().__init__(master, fg_color="transparent")
        self._order_id = order_id
        self._platform = platform
        self._neto_order = neto_order
        self._ebay_order = ebay_order
        self._neto_client = neto_client
        self._shipping_config = shipping_config
        self._dry_run = dry_run
        self._on_back = on_back
        self._on_courier_selected = on_courier_selected

        self._package_rows: list[PackageRow] = []
        self._courier_toggles: dict[str, ctk.BooleanVar] = {}
        self._courier_switches: dict[str, ctk.CTkSwitch] = {}
        self._courier_enabled_config: dict[str, bool] = {}
        self._quote_cards: list[QuoteCard] = []
        self._selected_quote: Quote | None = None
        self._last_request: ShipmentRequest | None = None
        self._couriers_by_code: dict = {}
        self._dims_auto_filled: bool = False  # True if Neto already had dimensions

        # Determine receiver address from order data
        self._single_sku: str | None = None  # Set if single line, qty 1
        if neto_order:
            self._receiver = address_from_neto_order(neto_order)
            self._shipping_type = neto_order.shipping_type or "Standard"
            self._order_value = neto_order.grand_total
            self._line_skus = [li.sku for li in neto_order.line_items]
            if len(neto_order.line_items) == 1 and neto_order.line_items[0].quantity == 1:
                self._single_sku = neto_order.line_items[0].sku
        elif ebay_order:
            self._receiver = address_from_ebay_order(ebay_order)
            self._shipping_type = ebay_order.shipping_type or "Standard"
            self._order_value = ebay_order.order_total
            self._line_skus = [li.sku for li in ebay_order.line_items]
            if len(ebay_order.line_items) == 1 and ebay_order.line_items[0].quantity == 1:
                self._single_sku = ebay_order.line_items[0].sku
        else:
            self._receiver = Address("", "", "", "", "", "", "", "AU")
            self._shipping_type = "Standard"
            self._order_value = 0.0
            self._line_skus = []

        self._sender = sender_from_config(shipping_config.sender)

        self._build_ui()
        self._auto_fill_dimensions()

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Nav bar
        nav = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), corner_radius=0)
        nav.grid(row=0, column=0, sticky="ew")

        ctk.CTkButton(
            nav, text="← Back to Order", width=150, height=32,
            fg_color="transparent", hover_color=("gray75", "gray30"),
            font=ctk.CTkFont(size=13), command=self._on_back,
        ).pack(side="left", padx=8, pady=6)

        ctk.CTkLabel(
            nav, text=f"Book Freight  —  Order {self._order_id}",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=8)

        # Scrollable body
        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew")

        self._build_address_section(body)
        self._build_packages_section(body)
        self._build_sku_search(body)
        self._build_courier_section(body)
        self._build_action_section(body)
        self._build_results_section(body)

    # ── Section 1: Shipping Address ──────────────────────────────────────

    def _build_address_section(self, parent):
        section = ctk.CTkFrame(parent, border_width=1, border_color=("gray70", "gray35"))
        section.pack(fill="x", padx=10, pady=(10, 6))

        ctk.CTkLabel(section, text="Shipping Address", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        # Shipping type toggle
        type_row = ctk.CTkFrame(section, fg_color="transparent")
        type_row.pack(fill="x", padx=10, pady=(0, 4))
        ctk.CTkLabel(type_row, text="Shipping Type:", font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        # Default to "Standard" unless the order's shipping type is "Express"
        default_type = "Express" if self._shipping_type == "Express" else "Standard"
        self._shipping_type_var = ctk.StringVar(value=default_type)
        for val in ("Standard", "Express"):
            ctk.CTkRadioButton(
                type_row, text=val, variable=self._shipping_type_var, value=val,
                font=ctk.CTkFont(size=11),
                command=self._on_shipping_type_changed,
            ).pack(side="left", padx=(0, 12))

        # Two-column grid for address fields
        grid = ctk.CTkFrame(section, fg_color="transparent")
        grid.pack(fill="x", padx=10, pady=(0, 8))
        grid.grid_columnconfigure((0, 1), weight=1)

        self._addr_entries = {}
        fields_left = [
            ("Name", self._receiver.name),
            ("Company", self._receiver.company),
            ("Street 1", self._receiver.street1),
            ("Street 2", self._receiver.street2),
        ]
        fields_right = [
            ("Suburb", self._receiver.city),
            ("State", self._receiver.state),
            ("Postcode", self._receiver.postcode),
            ("Phone", self._receiver.phone),
        ]

        for i, (label, val) in enumerate(fields_left):
            self._addr_field(grid, label, val, row=i, col=0)
        for i, (label, val) in enumerate(fields_right):
            self._addr_field(grid, label, val, row=i, col=1)

    def _addr_field(self, parent, label: str, value: str, row: int, col: int):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=col, sticky="ew", padx=4, pady=2)
        ctk.CTkLabel(frame, text=label + ":", font=ctk.CTkFont(size=10), width=60, anchor="e").pack(
            side="left", padx=(0, 4)
        )
        entry = ctk.CTkEntry(frame, font=ctk.CTkFont(size=11))
        entry.pack(side="left", fill="x", expand=True)
        entry.insert(0, value)
        self._addr_entries[label] = entry

    def _get_receiver_address(self) -> Address:
        return Address(
            name=self._addr_entries["Name"].get(),
            company=self._addr_entries["Company"].get(),
            street1=self._addr_entries["Street 1"].get(),
            street2=self._addr_entries["Street 2"].get(),
            city=self._addr_entries["Suburb"].get(),
            state=self._addr_entries["State"].get(),
            postcode=self._addr_entries["Postcode"].get(),
            country="AU",
            phone=self._addr_entries["Phone"].get(),
            email=self._receiver.email,
        )

    # ── Section 2: Packages ──────────────────────────────────────────────

    def _build_packages_section(self, parent):
        section = ctk.CTkFrame(parent, border_width=1, border_color=("gray70", "gray35"))
        section.pack(fill="x", padx=10, pady=6)

        header = ctk.CTkFrame(section, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(header, text="Packages", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(
            header, text="+ Add Package", width=110, height=26,
            font=ctk.CTkFont(size=11), command=self._add_package_row,
        ).pack(side="right")

        self._packages_container = ctk.CTkFrame(section, fg_color="transparent")
        self._packages_container.pack(fill="x", padx=10, pady=(0, 8))

        # Create one row per line item SKU
        if self._line_skus:
            for sku in self._line_skus:
                self._add_package_row(sku=sku)
        else:
            self._add_package_row()

        # "Save dimensions to Neto" checkbox — only for single-line, qty-1 orders
        self._save_dims_var = ctk.BooleanVar(value=False)
        self._save_dims_check = None
        if self._single_sku and self._neto_client and self._platform.lower() == "neto":
            check_frame = ctk.CTkFrame(section, fg_color="transparent")
            check_frame.pack(fill="x", padx=10, pady=(0, 8))
            self._save_dims_check = ctk.CTkCheckBox(
                check_frame,
                text=f"Save dimensions to Neto for SKU: {self._single_sku}",
                variable=self._save_dims_var,
                font=ctk.CTkFont(size=11),
            )
            self._save_dims_check.pack(side="left")

    def _add_package_row(self, sku: str = ""):
        row = PackageRow(
            self._packages_container, sku=sku,
            on_remove=lambda: self._remove_package_row(row),
        )
        row.pack(fill="x", pady=3)
        self._package_rows.append(row)

    def _remove_package_row(self, row: PackageRow):
        if len(self._package_rows) <= 1:
            return  # Keep at least one
        row.destroy()
        self._package_rows.remove(row)

    # ── SKU Search ───────────────────────────────────────────────────────

    def _build_sku_search(self, parent):
        section = ctk.CTkFrame(parent, fg_color="transparent")
        section.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(section, text="Search SKU for dimensions:", font=ctk.CTkFont(size=11)).pack(
            side="left", padx=(0, 6)
        )
        self._sku_search_entry = ctk.CTkEntry(section, width=150, font=ctk.CTkFont(size=11),
                                               placeholder_text="Enter SKU...")
        self._sku_search_entry.pack(side="left", padx=(0, 6))
        self._sku_search_btn = ctk.CTkButton(
            section, text="Search", width=70, height=26,
            font=ctk.CTkFont(size=11), command=self._on_sku_search,
        )
        self._sku_search_btn.pack(side="left", padx=(0, 6))
        self._sku_search_status = ctk.CTkLabel(section, text="", font=ctk.CTkFont(size=10), text_color="gray50")
        self._sku_search_status.pack(side="left")

    def _on_sku_search(self):
        sku = self._sku_search_entry.get().strip()
        if not sku or not self._neto_client:
            return
        self._sku_search_btn.configure(state="disabled", text="...")
        self._sku_search_status.configure(text="Searching...")

        def _fetch():
            dims = self._neto_client.get_item_dimensions(sku)
            self.after(0, lambda: self._on_sku_search_result(sku, dims))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_sku_search_result(self, sku: str, dims: dict | None):
        self._sku_search_btn.configure(state="normal", text="Search")
        if dims:
            # Apply to the first package row (or selected row)
            if self._package_rows:
                self._package_rows[0].set_from_dimensions(dims, source=f"SKU {sku}")
            self._sku_search_status.configure(text=f"Found: {sku}", text_color="green")
        else:
            self._sku_search_status.configure(text=f"No dimensions for {sku}", text_color="orange")

    # ── Section 3: Courier Selection ─────────────────────────────────────

    def _build_courier_section(self, parent):
        section = ctk.CTkFrame(parent, border_width=1, border_color=("gray70", "gray35"))
        section.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(section, text="Couriers", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        toggles_frame = ctk.CTkFrame(section, fg_color="transparent")
        toggles_frame.pack(fill="x", padx=10, pady=(0, 8))

        courier_configs = self._shipping_config.couriers
        courier_names = {
            "auspost": "Australia Post",
            "aramex": "Aramex",
            "bonds": "Bonds Couriers",
            "allied": "Allied Express",
            "dai_post": "DAI Post",
        }

        self._courier_switches: dict[str, ctk.CTkSwitch] = {}
        for code, display_name in courier_names.items():
            cfg = courier_configs.get(code, {})
            enabled = cfg.get("enabled", False)
            var = ctk.BooleanVar(value=enabled)
            switch = ctk.CTkSwitch(
                toggles_frame, text=display_name, variable=var,
                font=ctk.CTkFont(size=11),
            )
            if not enabled:
                switch.configure(state="disabled")
            switch.pack(side="left", padx=(0, 16))
            self._courier_toggles[code] = var
            self._courier_switches[code] = switch
            self._courier_enabled_config[code] = enabled

        # Apply express restrictions if the order defaults to express
        self._on_shipping_type_changed()

    def _on_shipping_type_changed(self):
        """Enable/disable courier switches based on the selected shipping type."""
        is_express = self._shipping_type_var.get() == "Express"
        for code, switch in self._courier_switches.items():
            configured = self._courier_enabled_config.get(code, False)
            if not configured:
                continue  # Already disabled, leave it
            if is_express and code not in EXPRESS_CAPABLE_COURIERS:
                switch.deselect()
                switch.configure(state="disabled")
            else:
                switch.configure(state="normal")

    # ── Section 4: Action ────────────────────────────────────────────────

    def _build_action_section(self, parent):
        section = ctk.CTkFrame(parent, fg_color="transparent")
        section.pack(fill="x", padx=10, pady=6)

        self._quote_btn = ctk.CTkButton(
            section, text="Get Quotes", width=160, height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
            command=self._on_get_quotes,
        )
        self._quote_btn.pack(side="left", padx=(0, 12))

        self._progress_label = ctk.CTkLabel(section, text="", font=ctk.CTkFont(size=11))
        self._progress_label.pack(side="left")

    # ── Section 5: Quote Results ─────────────────────────────────────────

    def _build_results_section(self, parent):
        self._results_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._results_frame.pack(fill="x", padx=10, pady=6)
        # Hidden until quotes arrive

        self._use_btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._use_btn_frame.pack(fill="x", padx=10, pady=(0, 12))

    # ── Auto-fill dimensions from Neto ───────────────────────────────────

    def _auto_fill_dimensions(self):
        if not self._neto_client or not self._line_skus:
            return
        for i, sku in enumerate(self._line_skus):
            if i < len(self._package_rows):
                row = self._package_rows[i]
                self._fetch_dimensions_for_row(sku, row)

    def _fetch_dimensions_for_row(self, sku: str, row: PackageRow):
        def _fetch():
            dims = self._neto_client.get_item_dimensions(sku)
            if dims:
                def _apply():
                    row.set_from_dimensions(dims, source="auto-filled from Neto")
                    self._dims_auto_filled = True
                    # Dimensions already exist — change checkbox to "overwrite" mode
                    if self._save_dims_check is not None:
                        self._save_dims_check.configure(
                            text=f"Overwrite dimensions on Neto for SKU: {self._single_sku}"
                        )
                        self._save_dims_var.set(False)
                self.after(0, _apply)

        threading.Thread(target=_fetch, daemon=True).start()

    # ── Get Quotes ───────────────────────────────────────────────────────

    def _on_get_quotes(self):
        # Validate packages
        packages = []
        for row in self._package_rows:
            pkg = row.get_package()
            if pkg is None or pkg.weight_kg <= 0:
                self._progress_label.configure(text="Please fill in all package dimensions", text_color="red")
                return
            packages.append(pkg)

        # Build request
        receiver = self._get_receiver_address()
        shipping_type = self._shipping_type_var.get()

        request = ShipmentRequest(
            order_id=self._order_id,
            platform=self._platform,
            sender=self._sender,
            receiver=receiver,
            packages=packages,
            shipping_type=shipping_type,
            order_value=self._order_value,
            dry_run=self._dry_run,
        )

        # Determine enabled couriers
        enabled = {code for code, var in self._courier_toggles.items() if var.get()}
        if not enabled:
            self._progress_label.configure(text="No couriers enabled", text_color="red")
            return

        self._quote_btn.configure(state="disabled", text="Quoting...")
        self._progress_label.configure(text="Fetching quotes...", text_color=("gray20", "gray80"))
        self._clear_results()

        # Build couriers and engine — store instances for later booking
        couriers = _build_couriers(self._shipping_config.couriers)
        self._couriers_by_code = {c.code: c for c in couriers}
        self._last_request = request
        engine = QuoteEngine(couriers)

        def _progress(name: str, status: str):
            self.after(0, lambda: self._progress_label.configure(
                text=f"{name}: {status}..."
            ))

        def _run():
            quotes = engine.get_quotes(request, enabled_codes=enabled, progress_callback=_progress)
            self.after(0, lambda: self._show_results(quotes))

        threading.Thread(target=_run, daemon=True).start()

    def _clear_results(self):
        for card in self._quote_cards:
            card.destroy()
        self._quote_cards.clear()
        self._selected_quote = None
        for w in self._use_btn_frame.winfo_children():
            w.destroy()

    def _show_results(self, quotes: list[Quote]):
        # Quotes received — lock Get Quotes button (turn blue) so user is
        # directed toward "Use Selected Courier" instead
        self._quote_btn.configure(
            state="disabled", text="Get Quotes",
            fg_color=("dodgerblue", "dodgerblue3"),
            hover_color=("dodgerblue3", "dodgerblue4"),
        )

        if not quotes:
            self._progress_label.configure(text="No quotes returned", text_color="orange")
            return

        successful = [q for q in quotes if not q.error]
        cheapest_price = min((q.price for q in successful), default=None)
        count = len(successful)
        self._progress_label.configure(
            text=f"{count} quote{'s' if count != 1 else ''} received",
            text_color="green" if count else "orange",
        )

        for quote in quotes:
            is_cheapest = (not quote.error and quote.price == cheapest_price and cheapest_price is not None)
            card = QuoteCard(
                self._results_frame, quote=quote,
                on_select=self._select_quote, is_cheapest=is_cheapest,
            )
            card.pack(fill="x", pady=3)
            self._quote_cards.append(card)

        # "Use Selected Courier" button — starts disabled/neutral; turns green when a quote is selected
        self._use_selected_btn = ctk.CTkButton(
            self._use_btn_frame, text="Use Selected Courier", width=180, height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("gray50", "gray40"),
            state="disabled",
            command=self._use_selected,
        )
        self._use_selected_btn.pack(side="left", pady=6)

    def _select_quote(self, quote: Quote):
        self._selected_quote = quote
        for card in self._quote_cards:
            card.set_selected(card._quote is quote)
        if hasattr(self, "_use_selected_btn"):
            self._use_selected_btn.configure(
                state="normal",
                fg_color=("green3", "green4"),
                hover_color=("green4", "green3"),
            )

    def _upload_dimensions_if_checked(self):
        """Upload package dimensions to Neto if the checkbox is checked."""
        if not self._save_dims_var.get() or not self._single_sku or not self._neto_client:
            return
        if not self._package_rows:
            return
        pkg = self._package_rows[0].get_package()
        if pkg is None:
            return

        sku = self._single_sku
        log.info("Uploading dimensions for SKU %s: %.2fkg  %.1fx%.1fx%.1fcm",
                 sku, pkg.weight_kg, pkg.length_cm, pkg.width_cm, pkg.height_cm)

        def _upload():
            try:
                self._neto_client.update_item_dimensions(
                    sku=sku,
                    weight_kg=pkg.weight_kg,
                    length_cm=pkg.length_cm,
                    width_cm=pkg.width_cm,
                    height_cm=pkg.height_cm,
                    dry_run=self._dry_run,
                )
                log.info("Dimensions uploaded for SKU %s", sku)
            except Exception as exc:
                log.error("Failed to upload dimensions for SKU %s: %s", sku, exc)

        threading.Thread(target=_upload, daemon=True).start()

    def _use_selected(self):
        if not self._selected_quote or not self._last_request:
            return

        # Upload dimensions to Neto if checkbox is checked
        self._upload_dimensions_if_checked()

        quote = self._selected_quote
        log.info("Courier selected: %s (%s)  price=$%.2f",
                 quote.courier_name, quote.courier_code, quote.price)

        # Dry-run: skip real booking
        if self._dry_run:
            log.info("Dry-run mode — skipping real booking")
            self._on_courier_selected(quote.courier_name, "DRY-RUN-TRACKING")
            return

        courier = self._couriers_by_code.get(quote.courier_code)
        if courier is None:
            log.warning("Courier code '%s' not found in couriers_by_code — manual tracking only",
                        quote.courier_code)
            self._on_courier_selected(quote.courier_name, "")
            return

        log.info("Starting booking thread for courier '%s', order '%s'",
                 quote.courier_code, self._last_request.order_id)
        self._use_selected_btn.configure(state="disabled", text="Booking...")
        self._progress_label.configure(text="Confirming booking with courier...",
                                        text_color=("gray20", "gray80"))
        request = self._last_request

        def _do_book():
            log.debug("Calling %s.book() for order %s", courier.code, request.order_id)
            result = courier.book(request, quote)
            log.debug("book() returned: tracking=%r  reference=%r  error=%r",
                      result.tracking_number, result.booking_reference, result.error or None)
            self.after(0, lambda: self._on_booking_result(result))

        threading.Thread(target=_do_book, daemon=True).start()

    def _on_booking_result(self, result):
        self._use_selected_btn.configure(state="normal", text="Use Selected Courier")

        if result.error:
            if "not yet implemented" in result.error.lower():
                log.info("'%s' has no booking API — returning for manual tracking entry",
                         result.courier_name)
                self._on_courier_selected(result.courier_name, "")
            else:
                log.error("Booking failed for '%s': %s", result.courier_name, result.error)
                self._progress_label.configure(
                    text=f"Booking failed: {result.error}", text_color="red"
                )
            return

        log.info("Booking confirmed — courier=%s  tracking=%s  reference=%s",
                 result.courier_name, result.tracking_number, result.booking_reference)

        # Resolve courier_code for ledger, capture, and printing
        courier_code = ""
        for code, c in self._couriers_by_code.items():
            if c.name == result.courier_name:
                courier_code = code
                break

        # Record booking in daily ledger
        bookings_dir = self._shipping_config.bookings_dir
        if bookings_dir and result.tracking_number:
            try:
                from src.shipping.booking_ledger import add_booking
                extras = {}
                if result.booking_reference:
                    extras["booking_reference"] = result.booking_reference
                if self._receiver.postcode:
                    extras["postcode"] = self._receiver.postcode
                add_booking(
                    directory=bookings_dir,
                    courier_code=courier_code,
                    courier_name=result.courier_name,
                    tracking_number=result.tracking_number,
                    order_id=self._order_id,
                    recipient=self._receiver.name,
                    extras=extras if extras else None,
                )
            except Exception as exc:
                log.warning("Failed to record booking in ledger: %s", exc)

        # AusPost Express has a vertical barcode spanning the full label height — print as single strip
        _no_split = (
            courier_code == "auspost"
            and self._last_request is not None
            and self._last_request.shipping_type == "Express"
        )
        _print_courier_code = "auspost_express" if (
            courier_code == "auspost" and _no_split
        ) else courier_code

        # Always save the latest label for this courier (overwrites previous)
        if result.label_pdf and _print_courier_code:
            try:
                from src.shipping.label_capture import save_label
                save_label(_print_courier_code, result.label_pdf)
            except Exception as exc:
                log.warning("Label save failed: %s", exc)

        # Print label in background thread (non-fatal if it fails).
        # Navigation happens immediately below (freight view is destroyed), so we
        # show any error via the root window which outlives this widget.
        if result.label_pdf:
            log.debug("Label PDF received (%d bytes) — starting print thread", len(result.label_pdf))
            label_bytes = result.label_pdf
            root = self.winfo_toplevel()

            def _print():
                from src.shipping.label_printer import print_label
                from tkinter import messagebox
                err = print_label(label_bytes, courier_code=_print_courier_code, no_split=_no_split)
                if err:
                    log.error("Label print failed: %s", err)
                    root.after(0, lambda: messagebox.showwarning(
                        "Label Print Failed",
                        f"The booking was confirmed but the label could not be printed:\n\n{err}\n\n"
                        "You can re-print manually from the Aramex portal.",
                        parent=root,
                    ))
                else:
                    log.info("Label printed successfully")

            threading.Thread(target=_print, daemon=True).start()
        else:
            log.warning("No label PDF in booking result — nothing to print")

        self._on_courier_selected(result.courier_name, result.tracking_number)
