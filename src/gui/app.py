import customtkinter as ctk

from src.config import ConfigManager
from src.ebay_client import EbayClient
from src.neto_client import NetoClient


class App(ctk.CTk):
    def __init__(self, config: ConfigManager):
        super().__init__()
        self.config = config

        self.title("Scarlett Music — Overdue Orders Matcher")
        self.geometry("1150x720")
        self.minsize(900, 600)

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

        self._build_ui()

    def _build_ui(self):
        # Header bar
        header = ctk.CTkFrame(self, height=48, corner_radius=0, fg_color=("gray85", "gray20"))
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header,
            text="Scarlett Music  —  Overdue Orders Matcher",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(side="left", padx=20, pady=10)

        # Tab view
        self.tabview = ctk.CTkTabview(self, corner_radius=8)
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
