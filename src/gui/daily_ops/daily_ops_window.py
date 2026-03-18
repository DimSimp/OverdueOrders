from __future__ import annotations

import os

import customtkinter as ctk

from src.version import __version__

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
_APP_ICON = os.path.join(_ROOT, "AIO.ico")


class DailyOpsWindow(ctk.CTkToplevel):
    """
    Daily Operations main window.

    Opens on a menu screen with three top-level actions. Each action uses a
    stacked-frame navigation model — frames are placed in the same grid cell
    and raised to the top via _show_step().

    Picking list workflow steps:
        1. OptionsView   — platform/filter toggles + date range
        2. FetchView     — background order fetch with progress
        3. EnvelopeView  — classify + print envelope orders   (Phase 2b)
        4. PickZoneView  — assign pick zones to items          (Phase 2c)
        5. PickListView  — preview + export picking list CSV   (Phase 2c)
        6. ResultsView   — order list + freight booking        (Phase 2d)
    """

    def __init__(self, master, config, neto_client, ebay_client):
        super().__init__(master)
        self.config = config
        self.neto_client = neto_client
        self.ebay_client = ebay_client

        # ── Shared state set by fetch step ──────────────────────────────
        self.neto_orders: list = []
        self.ebay_orders: list = []
        # Set after envelope classification (Phase 2b)
        self.envelope_classifications: dict = {}   # order_id → "minilope"/"devilope"/"satchel"
        # Set after pick zone assignment (Phase 2c)
        self.pick_zones: dict = {}                 # sku → zone_name

        self.title("Scarlett Music — Daily Operations")
        self.geometry("1150x720")
        self.minsize(900, 600)
        if os.path.exists(_APP_ICON):
            try:
                self.iconbitmap(_APP_ICON)
            except Exception:
                pass

        self._build_ui()
        # Bring to front
        self.lift()
        self.focus_force()

    # ── Layout ──────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, height=48, corner_radius=0, fg_color=("gray85", "gray20"))
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        self._header_label = ctk.CTkLabel(
            header,
            text="Daily Operations",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self._header_label.pack(side="left", padx=20, pady=10)

        ctk.CTkLabel(
            header,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=16, pady=10)

        # Step indicator (right side of header)
        self._step_label = ctk.CTkLabel(
            header,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        )
        self._step_label.pack(side="right", padx=8)

        # Dry-run banner
        if self.config.app.dry_run:
            dry_banner = ctk.CTkFrame(self, height=28, corner_radius=0, fg_color=("red3", "red4"))
            dry_banner.pack(fill="x", side="top")
            dry_banner.pack_propagate(False)
            ctk.CTkLabel(
                dry_banner,
                text="DRY RUN MODE — API writes are simulated",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="white",
            ).pack(pady=4)

        # Content area — all step frames are stacked here
        self._content = ctk.CTkFrame(self, fg_color="transparent")
        self._content.pack(fill="both", expand=True)
        self._content.grid_rowconfigure(0, weight=1)
        self._content.grid_columnconfigure(0, weight=1)

        self._step_frames: dict = {}

        # Start at the menu
        self._show_menu()

    # ── Step navigation ──────────────────────────────────────────────────

    def _show_step(self, frame: ctk.CTkFrame, step_text: str = ""):
        frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()
        self._step_label.configure(text=step_text)

    def set_header(self, title: str):
        """Allow step views to update the header title."""
        self._header_label.configure(text=title)

    # ── Menu ─────────────────────────────────────────────────────────────

    def _show_menu(self):
        if "menu" not in self._step_frames:
            self._step_frames["menu"] = _DailyOpsMenuView(
                self._content,
                on_picking_list=self._show_options,
                on_search_order=lambda: self._show_placeholder(
                    "Search for Order", "", "(Coming soon)"),
                on_show_orders=lambda: self._show_placeholder(
                    "Show All Orders", "", "(Coming soon)"),
            )
        self.set_header("Daily Operations")
        self._show_step(self._step_frames["menu"], "")

    # ── Step 1: Options ──────────────────────────────────────────────────

    def _show_options(self):
        if "options" not in self._step_frames:
            from src.gui.daily_ops.options_view import OptionsView
            self._step_frames["options"] = OptionsView(
                self._content,
                window=self,
                on_generate=self._start_fetch,
                on_back=self._show_menu,
            )
        self.set_header("Daily Operations  —  Picking List Options")
        self._show_step(self._step_frames["options"], "Step 1 of 6")

    # ── Step 2: Fetch ────────────────────────────────────────────────────

    def _start_fetch(self, options: dict):
        if "fetch" not in self._step_frames:
            from src.gui.daily_ops.fetch_view import FetchView
            self._step_frames["fetch"] = FetchView(
                self._content,
                window=self,
                on_complete=self._on_fetch_complete,
                on_back=self._show_options,
            )
        self.set_header("Daily Operations  —  Fetching Orders")
        self._show_step(self._step_frames["fetch"], "Step 2 of 6")
        self._step_frames["fetch"].start_fetch(options)

    def _on_fetch_complete(self):
        # Clear cached envelope/PDF frames so they're rebuilt with fresh order data
        for key in ("envelope_classify", "envelope_pdf"):
            if key in self._step_frames:
                self._step_frames.pop(key).destroy()
        self._show_envelope_classify()

    # ── Step 3: Envelope Classification ──────────────────────────────────

    def _show_envelope_classify(self):
        if "envelope_classify" not in self._step_frames:
            from src.gui.daily_ops.envelope_view import EnvelopeClassifyView
            self._step_frames["envelope_classify"] = EnvelopeClassifyView(
                self._content,
                window=self,
                on_complete=self._on_envelope_classify_complete,
                on_back=self._show_options,
            )
        self.set_header("Daily Operations  —  Classify Envelopes")
        self._show_step(self._step_frames["envelope_classify"], "Step 3 of 6")
        self._step_frames["envelope_classify"].start_classify()

    def _on_envelope_classify_complete(self):
        self._show_envelope_pdf()

    # ── Step 4: Envelope PDFs ─────────────────────────────────────────────

    def _show_envelope_pdf(self):
        if "envelope_pdf" not in self._step_frames:
            from src.gui.daily_ops.envelope_view import EnvelopePDFView
            self._step_frames["envelope_pdf"] = EnvelopePDFView(
                self._content,
                window=self,
                on_complete=self._on_envelope_pdf_complete,
                on_back=self._show_envelope_classify,
            )
        self.set_header("Daily Operations  —  Envelope PDFs")
        self._show_step(self._step_frames["envelope_pdf"], "Step 4 of 6")
        self._step_frames["envelope_pdf"].generate_pdfs()

    def _on_envelope_pdf_complete(self):
        # Phase 2c placeholder
        self._show_placeholder("Pick Zone Classification", "Step 5 of 6",
                               "(Coming in Phase 2c)")

    # ── Placeholder for future steps ────────────────────────────────────

    def _show_placeholder(self, title: str, step_text: str, message: str, back=None):
        key = f"placeholder_{title}"
        if key not in self._step_frames:
            frame = _PlaceholderView(
                self._content,
                message=message,
                on_back=back if back is not None else self._show_menu,
            )
            self._step_frames[key] = frame
        self.set_header(f"Daily Operations  —  {title}")
        self._show_step(self._step_frames[key], step_text)


class _DailyOpsMenuView(ctk.CTkFrame):
    """Top-level menu for Daily Operations — shows the three available actions."""

    def __init__(self, master, on_picking_list, on_search_order, on_show_orders, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        center = ctk.CTkFrame(self, fg_color="transparent")
        center.pack(expand=True, fill="both", padx=80, pady=40)

        ctk.CTkLabel(
            center,
            text="What would you like to do?",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray70"),
        ).pack(pady=(0, 24))

        # Generate Picking List
        ctk.CTkButton(
            center,
            text="Generate Picking List",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            command=on_picking_list,
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            center,
            text="Fetch today's orders, classify envelopes, assign pick zones, export CSV",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(pady=(0, 20))

        # Search for Order
        ctk.CTkButton(
            center,
            text="Search for Order",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=on_search_order,
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            center,
            text="Look up a specific order by ID, customer name, or SKU",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(pady=(0, 20))

        # Show All Orders
        ctk.CTkButton(
            center,
            text="Show All Orders",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=on_show_orders,
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            center,
            text="Browse and manage all current orders with freight booking",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack()


class _PlaceholderView(ctk.CTkFrame):
    """Temporary placeholder for steps not yet implemented."""

    def __init__(self, master, message: str, on_back=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        ctk.CTkLabel(
            self,
            text=message,
            font=ctk.CTkFont(size=14),
            text_color=("gray50", "gray60"),
        ).pack(expand=True)

        if on_back:
            ctk.CTkButton(
                self,
                text="← Back",
                width=100,
                command=on_back,
            ).pack(pady=(0, 30))
