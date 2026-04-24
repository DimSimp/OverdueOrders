from __future__ import annotations

import os

import customtkinter as ctk

from src.version import __version__

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
_APP_ICON = os.path.join(_ROOT, "AIO.ico")


class PosWindow(ctk.CTkToplevel):
    """POS / Till main window.

    Opens maximised. Contains five tabs (Till, Inventory, Customers,
    Purchasing, Reporting). Only the Till tab is implemented initially;
    the others show a placeholder.
    """

    def __init__(self, master, config, current_user=None, on_switch_user=None):
        super().__init__(master)
        self.config = config
        self.current_user = current_user
        self._on_switch_user = on_switch_user

        self.title("Scarlett Music — POS / Till")
        self.geometry("1280x800")
        # Keep the POS window wide enough for the current 3-column customer detail layout.
        self.minsize(1220, 650)
        self.resizable(True, True)

        if os.path.exists(_APP_ICON):
            try:
                self.iconbitmap(_APP_ICON)
            except Exception:
                pass

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        # Maximise after the window is fully drawn
        def _raise():
            self.state("zoomed")
            self.lift()
            self.focus_force()
        self.after(50, _raise)

    # ── Layout ──────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, height=48, corner_radius=0,
                               fg_color=("gray85", "gray20"))
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="POS / Till",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(side="left", padx=20, pady=10)

        ctk.CTkLabel(
            header,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=16, pady=10)

        if self.current_user:
            _u = self.current_user
            _display = (
                f"{_u.get('first_name', '')} {_u.get('last_name', '')}".strip()
                or _u.get("username", "")
            )
            if self._on_switch_user:
                ctk.CTkButton(
                    header,
                    text="Switch User",
                    font=ctk.CTkFont(size=11),
                    height=26, width=96,
                    fg_color="transparent",
                    border_width=1,
                    border_color=("gray60", "gray45"),
                    text_color=("gray40", "gray70"),
                    hover_color=("gray85", "gray25"),
                    command=self._on_switch_user,
                ).pack(side="right", padx=(0, 8), pady=10)
            ctk.CTkLabel(
                header,
                text=f"\U0001f464 {_display}",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            ).pack(side="right", padx=(0, 4), pady=10)

        # Tab view — no extra padding so tabs sit flush against the header
        self._tabs = ctk.CTkTabview(self, anchor="nw")
        self._tabs.pack(fill="both", expand=True, padx=0, pady=0)
        tabs = self._tabs

        for name in ("Till", "Inventory", "Customers", "Purchasing", "Reporting"):
            tabs.add(name)

        # Till tab — real implementation
        from src.gui.pos.till_tab import TillTab
        self._till_tab = TillTab(tabs.tab("Till"), current_user=self.current_user)
        self._till_tab.pack(fill="both", expand=True)

        # Inventory tab — real implementation
        from src.gui.inventory.inventory_tab import InventoryTab
        self._inventory_tab = InventoryTab(tabs.tab("Inventory"),
                                           current_user=self.current_user)
        self._inventory_tab.pack(fill="both", expand=True)

        # Wire cross-tab callbacks
        self._inventory_tab.on_add_to_cart = self._till_tab.add_item_to_cart

        def _go_inventory(query: str, auto_select: bool = False):
            self._tabs.set("Inventory")
            self._inventory_tab.search_and_focus(query, auto_select=auto_select)

        self._till_tab.navigate_to_inventory = _go_inventory
        self._till_tab.refresh_inventory = self._inventory_tab.refresh_search

        def _go_customer(uuid: str):
            self._tabs.set("Customers")
            self._customer_tab.show_customer(uuid)

        self._till_tab.navigate_to_customer = _go_customer

        def _load_customer_in_till(customer: dict):
            self._tabs.set("Till")
            self._till_tab.load_customer_profile(customer)

        def _refund_in_till(tx: dict):
            self._tabs.set("Till")
            self._till_tab._load_refund_cart(tx)

        # Customers tab — real implementation
        from src.gui.customers.customer_tab import CustomerTab
        self._customer_tab = CustomerTab(
            tabs.tab("Customers"),
            current_user=self.current_user,
            on_load_in_till=_load_customer_in_till,
            on_refund_in_till=_refund_in_till,
        )
        self._customer_tab.pack(fill="both", expand=True)

        # Stub tabs
        for name in ("Purchasing", "Reporting"):
            ctk.CTkLabel(
                tabs.tab(name),
                text=f"{name} — coming soon",
                font=ctk.CTkFont(size=14),
                text_color=("gray50", "gray60"),
            ).pack(expand=True)
