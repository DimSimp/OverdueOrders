from __future__ import annotations

import os
import threading
import webbrowser

import customtkinter as ctk

from src.config import ConfigManager
from src.ebay_client import EbayClient
from src.neto_client import NetoClient
from src.version import __version__

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_APP_ICON = os.path.join(_ROOT, "AIO.ico")


class App(ctk.CTk):
    def __init__(self, config: ConfigManager, startup_session: str | None = None):
        super().__init__()
        self.config = config
        self._startup_session = startup_session

        self.title("Scarlett Music — Overdue Orders")
        if os.path.exists(_APP_ICON):
            self.iconbitmap(_APP_ICON)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Shared state passed between tabs
        self.invoice_items = []     # list[InvoiceItem]
        self.neto_orders = []       # list[NetoOrder]
        self.ebay_orders = []       # list[EbayOrder]
        self.matched_orders = []    # list[MatchedOrder]

        # API clients
        self.neto_client = NetoClient(config.neto)
        self.ebay_client = EbayClient(
            config.ebay,
            token_save_callback=config.save_ebay_tokens,
        )

        from src.sku_alias_manager import SkuAliasManager
        self.sku_alias_manager = SkuAliasManager(config.app.sku_aliases_file)

        if startup_session:
            # Direct .scar open — skip home screen, go straight to afternoon ops
            self._setup_afternoon_window()
            self._build_ui()
            self._start_update_check()
            self.after(200, lambda: self.invoice_tab.load_session_from_path(startup_session))
        else:
            # Show home screen first
            self._setup_home_window()
            self._build_home_screen()

    # ── Window size helpers ────────────────────────────────────────────────

    def _setup_home_window(self):
        self.geometry("500x400")
        self.resizable(False, False)

    def _setup_afternoon_window(self):
        self.title("Scarlett Music — Overdue Orders Matcher")
        self.geometry("1150x720")
        self.minsize(900, 600)
        self.resizable(True, True)

    # ── Home screen ───────────────────────────────────────────────────────

    def _build_home_screen(self):
        from src.gui.home_window import HomeFrame
        self._home_frame = HomeFrame(
            self,
            on_afternoon=self._enter_afternoon_mode,
            on_daily=self._enter_daily_mode,
        )
        self._home_frame.pack(fill="both", expand=True)

    def _enter_afternoon_mode(self):
        win = ctk.CTkToplevel(self)
        win.title("Scarlett Music — Overdue Orders Matcher")
        win.geometry("1150x720")
        win.minsize(900, 600)
        win.resizable(True, True)
        self._build_ui(container=win)
        self._start_update_check()

    def _enter_daily_mode(self):
        from src.gui.daily_ops.daily_ops_window import DailyOpsWindow
        DailyOpsWindow(
            master=self,
            config=self.config,
            neto_client=self.neto_client,
            ebay_client=self.ebay_client,
        )

    # ── Afternoon Operations UI ───────────────────────────────────────────

    def _build_ui(self, container=None):
        """Build the afternoon operations UI.

        container: the window/frame to pack widgets into (defaults to self,
        which is used for direct .scar-file launches; pass a CTkToplevel when
        opening from the home screen so the home screen stays visible).
        """
        if container is None:
            container = self

        # Header bar
        header = ctk.CTkFrame(container, height=48, corner_radius=0, fg_color=("gray85", "gray20"))
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        self._ui_header = header  # stored so _show_update_banner can reference it
        ctk.CTkLabel(
            header,
            text="Scarlett Music  —  Overdue Orders Matcher",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(
            header,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=16, pady=10)

        # Dry-run banner
        if self.config.app.dry_run:
            dry_banner = ctk.CTkFrame(container, height=28, corner_radius=0, fg_color=("red3", "red4"))
            dry_banner.pack(fill="x", side="top")
            dry_banner.pack_propagate(False)
            ctk.CTkLabel(
                dry_banner,
                text="DRY RUN MODE — API writes are simulated (change dry_run in config.json to disable)",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="white",
            ).pack(pady=4)

        # Tab view
        self.tabview = ctk.CTkTabview(container, corner_radius=8)
        self.tabview.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        for tab_name in ("1. Invoice", "2. Orders", "3. Results"):
            self.tabview.add(tab_name)

        # Import tabs lazily to avoid circular imports at module level
        from src.gui.invoice_tab import InvoiceTab
        from src.gui.orders_tab import OrdersTab
        from src.gui.results_tab import ResultsTab

        self.invoice_tab = InvoiceTab(
            self.tabview.tab("1. Invoice"),
            app=self,
            on_complete=lambda: self.tabview.set("2. Orders"),
        )
        self.invoice_tab.pack(fill="both", expand=True)

        self.orders_tab = OrdersTab(
            self.tabview.tab("2. Orders"),
            app=self,
            on_complete=self._activate_results,
        )
        self.orders_tab.pack(fill="both", expand=True)

        self.results_tab = ResultsTab(
            self.tabview.tab("3. Results"),
            app=self,
        )
        self.results_tab.pack(fill="both", expand=True)

        self.tabview.set("1. Invoice")

    def _activate_results(self):
        self.tabview.set("3. Results")
        # Defer loading so the tab renders first, then populates
        self.after(50, self.results_tab.load_results)

    # ── Auto-update ───────────────────────────────────────────────────────────

    def _start_update_check(self):
        """Kick off a background thread to check GitHub for a newer release."""
        def _check():
            from src.updater import check_for_update
            result = check_for_update(__version__)
            if result:
                version, page_url, download_url = result
                # Schedule UI update on the main thread
                self.after(0, lambda: self._show_update_banner(version, page_url, download_url))

        t = threading.Thread(target=_check, daemon=True)
        t.start()

    def _show_update_banner(self, version: str, page_url: str, download_url: str):
        """Show a slim banner below the header when a new version is available."""
        header = getattr(self, "_ui_header", None) or self.winfo_children()[0]
        banner = ctk.CTkFrame(header.master, height=30, corner_radius=0, fg_color=("#2a6099", "#1a4a77"))
        banner.pack(fill="x", side="top", after=header)
        banner.pack_propagate(False)

        inner = ctk.CTkFrame(banner, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            inner,
            text=f"Update available: v{version}  —  ",
            font=ctk.CTkFont(size=12),
            text_color="white",
        ).pack(side="left")

        # If the release has a zip asset, offer a direct download link
        if download_url:
            ctk.CTkButton(
                inner,
                text="Download update",
                font=ctk.CTkFont(size=12, underline=True),
                fg_color="transparent",
                hover_color=("#1a4a77", "#0d3055"),
                text_color=("#9ecfff", "#9ecfff"),
                border_width=0,
                height=20,
                command=lambda: webbrowser.open(download_url),
            ).pack(side="left")
            ctk.CTkLabel(
                inner, text="  |  ", font=ctk.CTkFont(size=12), text_color=("gray70", "gray60"),
            ).pack(side="left")

        ctk.CTkButton(
            inner,
            text="View release notes",
            font=ctk.CTkFont(size=12, underline=True),
            fg_color="transparent",
            hover_color=("#1a4a77", "#0d3055"),
            text_color=("#9ecfff", "#9ecfff"),
            border_width=0,
            height=20,
            command=lambda: webbrowser.open(page_url),
        ).pack(side="left")
