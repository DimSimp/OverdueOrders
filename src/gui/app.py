from __future__ import annotations

import os
import threading
import webbrowser

import customtkinter as ctk

from src.config import ConfigManager
from src.ebay_client import EbayClient
from src.gui.styles import apply_styles
from src.neto_client import NetoClient
from src.version import __version__

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_APP_ICON = os.path.join(_ROOT, "AIO.ico")


class App(ctk.CTk):
    def __init__(self, config: ConfigManager, startup_session: str | None = None):
        super().__init__()
        apply_styles()
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

        from src.ebay_variation_sku_manager import EbayVariationSkuManager
        self.ebay_variation_manager = EbayVariationSkuManager(config.app.ebay_variation_skus_file)
        self.ebay_client = EbayClient(
            config.ebay,
            token_save_callback=config.save_ebay_tokens,
            variation_sku_manager=self.ebay_variation_manager,
        )

        from src.sku_alias_manager import SkuAliasManager
        self.sku_alias_manager = SkuAliasManager(config.app.sku_aliases_file)

        from src.musipos_client import MusiposClient
        _mc = config.musipos
        self.musipos_client = MusiposClient(_mc) if (_mc and _mc.enabled) else None

        # Supabase — initialize if credentials are present in config
        if config.supabase:
            from src.supabase_client import init_client as _supa_init
            _supa_init(config.supabase.url, config.supabase.key)

        # User system
        from src.user_manager import UserManager, USERS_FILE
        from src.assignment_manager import AssignmentManager
        self.user_manager = UserManager()
        self.assignment_manager = AssignmentManager()

        # Periodic backup of critical network files → local fallback copies
        from src.backup import schedule_backup
        from src.config import CONFIG_PATH, LOCAL_DATA_DIR
        schedule_backup(
            sources=[USERS_FILE, CONFIG_PATH],
            backup_dir=LOCAL_DATA_DIR / "backups",
        )

        self._current_user: dict | None = None          # set after login
        self._processing_order_id: str | None = None   # set when detail view is open
        self._order_close_callback = None               # registered by OrderDetailView

        if startup_session:
            # Direct .scar open — skip home screen, go straight to afternoon ops
            self._setup_afternoon_window()
            self._build_ui()
            self._start_update_check()
            self.after(200, lambda: self.invoice_tab.load_session_from_path(startup_session))
        else:
            # Show login / first-run screen
            self._setup_home_window()
            self._build_login_screen()
            self.protocol("WM_DELETE_WINDOW", self._logout_and_exit)

    # ── Window size helpers ────────────────────────────────────────────────

    def _setup_home_window(self):
        self.resizable(True, True)

    def _setup_afternoon_window(self):
        self.title("Scarlett Music — Overdue Orders Matcher")
        self.geometry("1150x720")
        self.minsize(900, 600)
        self.resizable(True, True)

    # ── Login / first-run screen ──────────────────────────────────────────

    def _build_login_screen(self):
        """Show login or first-run frame depending on whether users exist."""
        # Clear any existing frame
        for child in self.winfo_children():
            child.destroy()

        if not self.user_manager.is_available():
            self._show_server_unavailable()
            return

        if self.user_manager.has_any_users():
            self.geometry("500x430")
            self.resizable(False, False)
            from src.gui.login_frame import LoginFrame
            frame = LoginFrame(self, self.user_manager, self._on_login_success)
        else:
            self.geometry("500x560")
            self.resizable(True, True)
            from src.gui.first_run_frame import FirstRunFrame
            frame = FirstRunFrame(self, self.user_manager, self._on_login_success)
        frame.pack(fill="both", expand=True)

    def _show_server_unavailable(self):
        """Show a simple error screen when the network share is unreachable."""
        from src.version import __version__
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True)

        banner = ctk.CTkFrame(frame, height=70, corner_radius=0, fg_color=("gray85", "gray20"))
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)
        ctk.CTkLabel(
            banner,
            text="Scarlett Music  —  Overdue Orders",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left", padx=24, pady=10)
        ctk.CTkLabel(
            banner,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=20)

        body = ctk.CTkFrame(frame, fg_color="transparent")
        body.pack(expand=True)
        ctk.CTkLabel(
            body,
            text="Cannot reach the server.",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("red3", "#FF6B6B"),
        ).pack(pady=(0, 8))
        ctk.CTkLabel(
            body,
            text="Make sure you are connected to the network\nand the server is online.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray65"),
        ).pack(pady=(0, 16))
        ctk.CTkButton(
            body,
            text="Retry",
            width=120,
            command=self._build_login_screen,
        ).pack()

    def _on_login_success(self, user: dict):
        """Called by LoginFrame / FirstRunFrame after a successful login."""
        self._current_user = user
        self._start_heartbeat()
        self._build_home_screen()

    # ── Home screen ───────────────────────────────────────────────────────

    def _build_home_screen(self):
        for child in self.winfo_children():
            child.destroy()
        from src.gui.home_window import HomeFrame
        self._home_frame = HomeFrame(
            self,
            on_afternoon=self._enter_afternoon_mode,
            on_daily=self._enter_daily_mode,
            on_settings=self._open_settings,
            on_switch_user=self.switch_user,
            current_user=self._current_user,
            on_pos=self._enter_pos_mode,
            pos_dev_user=self.config.pos_dev_user,
        )
        self._home_frame.pack(fill="both", expand=True)
        self.after(10, lambda: self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}"))

    def _enter_pos_mode(self):
        from src.gui.pos.pos_window import PosWindow
        PosWindow(
            self,
            config=self.config,
            current_user=self._current_user,
            on_switch_user=self.switch_user,
        )

    def _open_settings(self):
        from src.gui.settings_window import SettingsWindow
        SettingsWindow(self, self.user_manager, self._current_user, on_switch_user=self.switch_user)

    def _enter_afternoon_mode(self):
        win = ctk.CTkToplevel(self)
        win.title("Scarlett Music — Overdue Orders Matcher")
        win.geometry("1150x720")
        win.minsize(900, 600)
        win.resizable(True, True)
        self._build_ui(container=win)
        self._start_update_check()
        win.after(50, lambda: win.state("zoomed"))
        win.protocol("WM_DELETE_WINDOW", lambda: self._on_workflow_window_close(win))
        win.after(250, self.lower)

    def _enter_daily_mode(self):
        from src.gui.daily_ops.daily_ops_window import DailyOpsWindow
        win = DailyOpsWindow(
            master=self,
            config=self.config,
            neto_client=self.neto_client,
            ebay_client=self.ebay_client,
            ebay_variation_manager=self.ebay_variation_manager,
            user_manager=self.user_manager,
            current_user=self._current_user,
            assignment_manager=self.assignment_manager,
            on_switch_user=self.switch_user,
        )
        win.after(250, self.lower)

    def _on_workflow_window_close(self, win):
        """Clear processing flag when a workflow window is closed, then destroy it."""
        if self._current_user and self._processing_order_id:
            try:
                self.user_manager.clear_processing_order(self._current_user["username"])
            except Exception:
                pass
            self._processing_order_id = None
            self._order_close_callback = None
        win.destroy()

    # ── Logout / switch user ─────────────────────────────────────────────────

    def switch_user(self):
        """Log out the current user and return to the login screen."""
        self._stop_heartbeat()
        if self._current_user:
            try:
                self.user_manager.logout(self._current_user["username"])
            except Exception:
                pass
        self._current_user = None
        self._processing_order_id = None
        self._order_close_callback = None
        self._build_login_screen()

    def _logout_and_exit(self):
        """Logout the current user and exit the application."""
        self._stop_heartbeat()
        if self._current_user:
            try:
                self.user_manager.logout(self._current_user["username"])
            except Exception:
                pass
        self.destroy()

    # ── Heartbeat & polling ──────────────────────────────────────────────────

    def _start_heartbeat(self):
        self._heartbeat_active = True
        self._heartbeat_thread = threading.Thread(
            target=self._heartbeat_loop, daemon=True, name="heartbeat"
        )
        self._heartbeat_thread.start()

    def _stop_heartbeat(self):
        self._heartbeat_active = False

    def _heartbeat_loop(self):
        import time
        while self._heartbeat_active:
            time.sleep(10)
            if not self._heartbeat_active or not self._current_user:
                break
            username = self._current_user["username"]
            try:
                self.user_manager.heartbeat(username)
            except Exception:
                pass
            # Check if we've been displaced from the order we're processing
            if self._processing_order_id:
                try:
                    session = self.user_manager.get_session(username)
                    if session.get("processing_order_id") != self._processing_order_id:
                        # Someone took over
                        taken_by = self.user_manager.get_processing_user(
                            self._processing_order_id
                        ) or "another user"
                        order_id = self._processing_order_id
                        self._processing_order_id = None
                        self.after(0, lambda b=taken_by, o=order_id: self._on_processing_taken_over(b, o))
                except Exception:
                    pass
            # Check if we've been force-logged-out from another device
            try:
                session = self.user_manager.get_session(username)
                if session and not session.get("logged_in"):
                    self._heartbeat_active = False
                    device = session.get("device_name", "another device")
                    self.after(0, lambda d=device: self._force_relogin(d))
            except Exception:
                pass

    def _on_processing_taken_over(self, taken_by: str, order_id: str):
        from tkinter import messagebox
        if self._order_close_callback:
            try:
                self._order_close_callback()
            except Exception:
                pass
            self._order_close_callback = None
        messagebox.showinfo(
            "Order Taken Over",
            f"{taken_by} has taken over order #{order_id}.\nYour processing session has been closed.",
            parent=self,
        )

    def _force_relogin(self, device: str):
        from tkinter import messagebox
        self._stop_heartbeat()
        self._current_user = None
        self._processing_order_id = None
        self._order_close_callback = None
        messagebox.showwarning(
            "Logged Out",
            f"You have been logged out because your account signed in on another device.",
            parent=self,
        )
        self._build_login_screen()

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
        _ver_label = ctk.CTkLabel(
            header,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        _ver_label.pack(side="right", padx=16, pady=10)
        _ver_label.bind("<Button-3>", lambda e: self._open_dev_console(container))
        if self._current_user:
            _u = self._current_user
            _display = (f"{_u.get('first_name', '')} {_u.get('last_name', '')}".strip()
                        or _u.get("username", ""))
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
                command=self.switch_user,
            ).pack(side="right", padx=(0, 8), pady=10)
            ctk.CTkLabel(
                header,
                text=f"👤 {_display}",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            ).pack(side="right", padx=(0, 4), pady=10)

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

        # Step container — frames are swapped in/out directly (no tabview)
        self._step_container = ctk.CTkFrame(container, fg_color="transparent")
        self._step_container.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        # Import lazily to avoid circular imports at module level
        from src.gui.invoice_tab import InvoiceTab
        from src.gui.results_tab import ResultsTab
        from src.gui.aft_menu_frame import AftMenuFrame

        # invoice_tab is never shown but holds invoice items for the matching engine
        self.invoice_tab = InvoiceTab(self._step_container, app=self, on_complete=lambda: None)
        self.results_tab = ResultsTab(self._step_container, app=self)

        self._aft_menu = AftMenuFrame(
            self._step_container,
            on_new_session=self._show_new_session,
            on_load_session=self._load_aft_session,
        )
        self._aft_menu.pack(fill="both", expand=True)

    def _hide_all_steps(self):
        for attr in ("_aft_menu", "_new_session_view", "invoice_tab", "results_tab"):
            f = getattr(self, attr, None)
            if f is not None:
                f.pack_forget()

    def _show_aft_menu(self):
        self._hide_all_steps()
        self._aft_menu.pack(fill="both", expand=True)

    def _show_new_session(self):
        from src.gui.aft_new_session_view import AftNewSessionView
        existing = getattr(self, "_new_session_view", None)
        if existing is None or not existing.winfo_exists():
            self._new_session_view = AftNewSessionView(
                self._step_container,
                app=self,
                on_complete=self._activate_results,
                on_back=self._show_aft_menu,
            )
        self._hide_all_steps()
        self._new_session_view.pack(fill="both", expand=True)

    def _load_aft_session(self):
        self.invoice_tab.load_session()

    def _show_results_tab(self):
        """Switch to results frame without re-running matching (used by session loads)."""
        self._hide_all_steps()
        self.results_tab.pack(fill="both", expand=True)

    def _activate_results(self):
        self._show_results_tab()
        # Defer loading so the frame renders first, then populates
        self.after(50, self.results_tab.load_results)

    def _open_dev_console(self, window=None):
        from src.gui.dev_console import open_dev_console
        open_dev_console(window or self)

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
