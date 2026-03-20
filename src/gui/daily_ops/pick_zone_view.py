from __future__ import annotations

import logging
import threading
from io import BytesIO

import customtkinter as ctk

log = logging.getLogger(__name__)

VALID_ZONES = ["String Room", "Back Area", "Out Front", "Picks"]

# Button colours for each zone (fg, hover) — None → default CTk blue
_ZONE_COLORS: list[tuple] = [
    (None, None),                                               # String Room — blue
    (("gray60", "gray40"), ("gray50", "gray35")),               # Back Area
    (("gray70", "gray30"), ("gray60", "gray25")),               # Out Front
    (("#2E7D32", "#1B5E20"), ("#256528", "#164A18")),            # Picks — green
]


class PickZoneView(ctk.CTkFrame):
    """
    Step 5 — Pick Zone Classification.

    Unique SKUs from all orders are checked against window.sku_attr_map
    (populated in Step 2).  SKUs that already have a valid zone are
    pre-classified automatically.  The rest are shown one at a time for
    manual assignment.

    Results stored in window.pick_zones: {sku: zone}.
    """

    def __init__(self, master, window, on_complete, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_complete = on_complete
        self._on_back = on_back
        self._queue: list[str] = []          # SKUs needing manual assignment
        self._current_idx: int = 0
        self._sku_info: dict[str, dict] = {} # sku → {name, order_count, current_zone}
        self._started = False
        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(18, 14))

        ctk.CTkLabel(
            header,
            text="Assign Pick Zones",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(side="left")

        self._progress_label = ctk.CTkLabel(
            header,
            text="",
            font=ctk.CTkFont(size=15),
            text_color=("gray50", "gray60"),
        )
        self._progress_label.pack(side="right")

        # ── Content area — stacked frames ─────────────────────────────────
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=24)
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)

        # Status frame (all done / no queue)
        self._status_frame = ctk.CTkFrame(content, fg_color="transparent")
        self._status_frame.grid(row=0, column=0, sticky="nsew")

        self._status_label = ctk.CTkLabel(
            self._status_frame,
            text="",
            font=ctk.CTkFont(size=16),
            text_color=("gray50", "gray60"),
        )
        self._status_label.pack(expand=True)

        # Classify frame (shown during interactive classification)
        self._classify_frame = ctk.CTkFrame(content, fg_color="transparent")
        self._classify_frame.grid(row=0, column=0, sticky="nsew")

        # SKU card
        card = ctk.CTkFrame(self._classify_frame, corner_radius=10)
        card.pack(fill="x", pady=(0, 16))

        card_inner = ctk.CTkFrame(card, fg_color="transparent")
        card_inner.pack(fill="x", padx=20, pady=20)

        self._img_label = ctk.CTkLabel(
            card_inner,
            text="",
            width=200,
            height=200,
            fg_color=("gray80", "gray30"),
            corner_radius=6,
        )
        self._img_label.pack(side="left", padx=(0, 24))

        # Persistent placeholder to prevent GC-triggered TclError
        from PIL import Image as _PILImage
        _ph = _PILImage.new("RGB", (200, 200), (180, 180, 180))
        self._placeholder_img = ctk.CTkImage(light_image=_ph, dark_image=_ph, size=(200, 200))
        self._displayed_img: ctk.CTkImage | None = None

        info = ctk.CTkFrame(card_inner, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True)

        self._sku_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=22, weight="bold"), anchor="w"
        )
        self._sku_label.pack(fill="x")

        self._desc_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=15), anchor="w",
            wraplength=620, justify="left",
        )
        self._desc_label.pack(fill="x", pady=(4, 0))

        self._orders_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=14),
            text_color=("gray50", "gray60"), anchor="w",
        )
        self._orders_label.pack(fill="x", pady=(8, 0))

        self._current_zone_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"), anchor="w",
        )
        self._current_zone_label.pack(fill="x", pady=(4, 0))

        # Zone buttons
        btn_row = ctk.CTkFrame(self._classify_frame, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 8))

        for i, zone in enumerate(VALID_ZONES):
            fg, hover = _ZONE_COLORS[i]
            btn_kwargs: dict = dict(
                text=zone,
                font=ctk.CTkFont(size=16, weight="bold"),
                height=72,
                command=lambda z=zone: self._on_assign(z),
            )
            if fg:
                btn_kwargs["fg_color"] = fg
                btn_kwargs["hover_color"] = hover
            padx = (0, 4) if i == 0 else (4, 0) if i == len(VALID_ZONES) - 1 else 4
            ctk.CTkButton(btn_row, **btn_kwargs).pack(
                side="left", expand=True, fill="x", padx=padx
            )

        ctk.CTkButton(
            self._classify_frame, text="Skip for now",
            width=140, height=34,
            fg_color="transparent", hover_color=("gray80", "gray25"),
            border_width=1, text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=13),
            command=self._on_skip,
        ).pack(anchor="w", pady=(0, 8))

        self._status_frame.tkraise()

        # ── Bottom nav ────────────────────────────────────────────────────
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", side="bottom", padx=24, pady=(8, 16))

        ctk.CTkButton(
            bottom, text="← Back",
            width=120, fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=self._on_back,
        ).pack(side="left")

        self._next_btn = ctk.CTkButton(
            bottom, text="Next: Picking List →",
            width=200, state="disabled",
            command=self._on_complete,
        )
        self._next_btn.pack(side="right")

    # ── Entry point ───────────────────────────────────────────────────────────

    def start_classify(self):
        """Called by DailyOpsWindow each time this step is shown."""
        if self._started:
            return
        self._started = True
        self._do_preclassify()

    # ── Classification logic ──────────────────────────────────────────────────

    def _do_preclassify(self):
        all_orders = self._window.neto_orders + self._window.ebay_orders
        attr_map = getattr(self._window, "sku_attr_map", {})

        # Build per-SKU info from order line items (name + order count)
        sku_info: dict[str, dict] = {}
        for order in all_orders:
            for li in order.line_items:
                sku = li.sku
                if not sku:
                    continue
                if sku not in sku_info:
                    name = (
                        getattr(li, "product_name", None)
                        or getattr(li, "title", None)
                        or ""
                    )
                    sku_info[sku] = {
                        "name": name,
                        "order_count": 0,
                        "current_zone": attr_map.get(sku, {}).get("pick_zone", ""),
                    }
                sku_info[sku]["order_count"] += 1

        self._sku_info = sku_info

        # Pre-classify valid zones; queue the rest
        self._queue = []
        pre_classified = 0
        for sku, info in sku_info.items():
            zone = info["current_zone"]
            if zone in VALID_ZONES:
                self._window.pick_zones[sku] = zone
                pre_classified += 1
            else:
                self._queue.append(sku)

        total = len(sku_info)
        if not self._queue:
            if total == 0:
                msg = "No SKUs found."
            else:
                msg = (
                    f"All {pre_classified} SKU{'s' if pre_classified != 1 else ''} "
                    f"already have pick zones assigned."
                )
            self._status_label.configure(text=msg)
            self._status_frame.tkraise()
            self._next_btn.configure(state="normal")
        else:
            self._current_idx = 0
            self._show_sku(self._current_idx)
            self._update_progress()
            self._classify_frame.tkraise()

    def _show_sku(self, idx: int):
        sku = self._queue[idx]
        info = self._sku_info.get(sku, {})

        self._sku_label.configure(text=sku)
        self._desc_label.configure(text=info.get("name", "") or "(no description)")

        count = info.get("order_count", 0)
        self._orders_label.configure(
            text=f"Appears in {count} order{'s' if count != 1 else ''}"
        )

        current = info.get("current_zone", "")
        self._current_zone_label.configure(
            text=f"Current zone: {current}" if current else "No zone assigned"
        )

        # Reset image, then load asynchronously
        self._displayed_img = None
        self._img_label.configure(image=self._placeholder_img, text="")
        self._img_generation = getattr(self, "_img_generation", 0) + 1
        gen = self._img_generation

        def _fetch(s=sku, g=gen):
            try:
                url_map = self._window.neto_client.get_product_images([s])
                url = url_map.get(s, "")
                if url:
                    self._load_image_async(url, g)
            except Exception as exc:
                log.debug("Image fetch failed for SKU %s: %s", s, exc)

        threading.Thread(target=_fetch, daemon=True).start()

    def _load_image_async(self, url: str, generation: int):
        def _load():
            try:
                import requests as req
                from PIL import Image
                resp = req.get(url, timeout=5)
                resp.raise_for_status()
                img = Image.open(BytesIO(resp.content)).convert("RGB")
                img.thumbnail((200, 200))
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(200, 200))
                def _apply(i=ctk_img, g=generation):
                    if getattr(self, "_img_generation", 0) == g:
                        self._displayed_img = i
                        self._img_label.configure(image=i, text="")
                self.after(0, _apply)
            except Exception as exc:
                log.debug("Image download failed gen=%d: %s", generation, exc)

        threading.Thread(target=_load, daemon=True).start()

    def _update_progress(self):
        remaining = len(self._queue) - self._current_idx
        self._progress_label.configure(
            text=f"{remaining} of {len(self._queue)} remaining"
        )

    def _on_assign(self, zone: str):
        sku = self._queue[self._current_idx]
        self._window.pick_zones[sku] = zone
        dry_run = self._window.config.app.dry_run
        threading.Thread(
            target=self._window.neto_client.update_item_pick_zone,
            args=(sku, zone),
            kwargs={"dry_run": dry_run},
            daemon=True,
        ).start()
        self._advance()

    def _on_skip(self):
        self._advance()

    def _advance(self):
        self._current_idx += 1
        if self._current_idx >= len(self._queue):
            assigned = len(self._window.pick_zones)
            skipped = len(self._queue) - assigned + sum(
                1 for s in self._queue if s in self._window.pick_zones
                and self._window.pick_zones[s] not in VALID_ZONES
            )
            self._status_label.configure(
                text=(
                    f"Pick zones complete — "
                    f"{assigned} SKU{'s' if assigned != 1 else ''} assigned."
                )
            )
            self._status_frame.tkraise()
            self._next_btn.configure(state="normal")
        else:
            self._show_sku(self._current_idx)
            self._update_progress()
