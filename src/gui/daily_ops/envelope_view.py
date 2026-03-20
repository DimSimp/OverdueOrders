from __future__ import annotations

import logging
import os
import threading
from datetime import date

import customtkinter as ctk

log = logging.getLogger(__name__)

from src.neto_client import NetoOrder
from src.ebay_client import EbayOrder

def _find_sumatra() -> str | None:
    """Return path to SumatraPDF.exe, or None if not found."""
    import shutil
    candidates = [
        shutil.which("SumatraPDF"),
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.expanduser(r"~\AppData\Local\SumatraPDF\SumatraPDF.exe"),
    ]
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None


# Map classification label → PostageType string sent to Neto
_POSTAGE_TYPE_MAP = {
    "minilope": "Minilope",
    "devilope": "Devilope",
    "satchel":  "Satchel",
}


# ── Step 3: Envelope Classification ──────────────────────────────────────────

class EnvelopeClassifyView(ctk.CTkFrame):
    """
    Step 3 — Envelope classification.

    Single-item orders are pre-classified using existing PostageType (Neto)
    or PrivateNotes keywords (eBay: "Mini." / "Devil.").
    Remaining single-item orders are shown one-at-a-time for manual
    classification.  Multi-item orders are not touched here.

    Results are stored in window.envelope_classifications:
        {order_id: "minilope" | "devilope" | "satchel"}
    """

    def __init__(self, master, window, on_complete, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_complete = on_complete
        self._on_back = on_back
        self._queue: list = []          # single-item orders needing manual classification
        self._current_idx: int = 0
        self._classifications: dict[str, str] = {}
        self._started = False
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header row ────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(18, 14))

        ctk.CTkLabel(
            header,
            text="Classify Envelopes",
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

        # Status frame (shown when nothing to classify, or all done)
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

        # Order card
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

        # Persistent placeholder — used instead of image=None to avoid tkinter GC crash
        # (setting image=None releases CTkImage; Tk then can't find the destroyed PhotoImage)
        from PIL import Image as _PILImage
        _ph = _PILImage.new("RGB", (200, 200), (180, 180, 180))
        self._placeholder_img = ctk.CTkImage(light_image=_ph, dark_image=_ph, size=(200, 200))
        self._displayed_img: ctk.CTkImage | None = None  # Hard ref prevents GC

        info = ctk.CTkFrame(card_inner, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True)

        self._customer_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=22, weight="bold"), anchor="w"
        )
        self._customer_label.pack(fill="x")

        self._order_info_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=14),
            text_color=("gray50", "gray60"), anchor="w",
        )
        self._order_info_label.pack(fill="x", pady=(2, 0))

        self._sku_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=15), anchor="w"
        )
        self._sku_label.pack(fill="x", pady=(6, 0))

        self._desc_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=15), anchor="w",
            wraplength=620, justify="left",
        )
        self._desc_label.pack(fill="x", pady=(4, 0))

        self._existing_label = ctk.CTkLabel(
            info, text="", font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"), anchor="w",
        )
        self._existing_label.pack(fill="x", pady=(8, 0))

        # Classification buttons
        btn_row = ctk.CTkFrame(self._classify_frame, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 8))

        self._mini_btn = ctk.CTkButton(
            btn_row, text="Minilope",
            font=ctk.CTkFont(size=16, weight="bold"), height=72,
            command=lambda: self._on_classify("minilope"),
        )
        self._mini_btn.pack(side="left", expand=True, fill="x", padx=(0, 4))

        self._devil_btn = ctk.CTkButton(
            btn_row, text="Devilope",
            font=ctk.CTkFont(size=16, weight="bold"), height=72,
            fg_color=("gray60", "gray40"), hover_color=("gray50", "gray35"),
            command=lambda: self._on_classify("devilope"),
        )
        self._devil_btn.pack(side="left", expand=True, fill="x", padx=4)

        self._neither_btn = ctk.CTkButton(
            btn_row, text="Neither  (Satchel)",
            font=ctk.CTkFont(size=16), height=72,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=lambda: self._on_classify("satchel"),
        )
        self._neither_btn.pack(side="left", expand=True, fill="x", padx=4)

        self._books_btn = ctk.CTkButton(
            btn_row, text="Books",
            font=ctk.CTkFont(size=16), height=72,
            fg_color=("#8B4513", "#5C2D0A"), hover_color=("#6B3410", "#4A2408"),
            command=lambda: self._on_classify("books"),
        )
        self._books_btn.pack(side="left", expand=True, fill="x", padx=(4, 0))

        ctk.CTkButton(
            self._classify_frame, text="Skip for now",
            width=140, height=34,
            fg_color="transparent", hover_color=("gray80", "gray25"),
            border_width=1, text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=13),
            command=self._on_skip,
        ).pack(anchor="w", pady=(0, 8))

        # Start status frame on top (until start_classify is called)
        self._status_frame.tkraise()

        # ── Bottom nav ────────────────────────────────────────────────────
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", side="bottom", padx=20, pady=(8, 16))

        ctk.CTkButton(
            bottom, text="← Back to Options",
            width=160, fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=self._on_back,
        ).pack(side="left")

        self._next_btn = ctk.CTkButton(
            bottom, text="Next: Envelope PDFs →",
            width=200, state="disabled",
            command=self._proceed,
        )
        self._next_btn.pack(side="right")

    # ── Entry point ───────────────────────────────────────────────────────

    def start_classify(self):
        """Called by DailyOpsWindow each time this step is shown."""
        if self._started:
            return
        self._started = True
        self._do_preclassify()

    # ── Classification logic ──────────────────────────────────────────────

    def _do_preclassify(self):
        all_orders = self._window.neto_orders + self._window.ebay_orders
        self._queue = []
        self._classifications = {}

        for order in all_orders:
            if len(order.line_items) != 1:
                continue  # multi-item orders → not envelope candidates
            label = self._try_auto_classify(order)
            if label:
                self._classifications[order.order_id] = label
            else:
                self._queue.append(order)

        self._window.envelope_classifications = dict(self._classifications)

        total_single = len(self._classifications) + len(self._queue)
        auto_count = len(self._classifications)

        if not self._queue:
            if total_single == 0:
                msg = "No single-item orders found — nothing to classify."
            else:
                msg = (
                    f"{auto_count} single-item order{'s' if auto_count != 1 else ''} "
                    f"auto-classified from existing PostageType / notes."
                )
            self._status_label.configure(text=msg)
            self._status_frame.tkraise()
            self._next_btn.configure(state="normal")
        else:
            self._current_idx = 0
            self._show_order(self._current_idx)
            self._update_progress()
            self._classify_frame.tkraise()

    def _try_auto_classify(self, order) -> str | None:
        if isinstance(order, NetoOrder):
            pt = (order.line_items[0].postage_type if order.line_items else "").lower()
            if "minilope" in pt:
                return "minilope"
            if "devilope" in pt:
                return "devilope"
            if "satchel" in pt:
                return "satchel"
        elif isinstance(order, EbayOrder):
            # Check Neto product catalogue first (persists across sessions)
            sku = order.line_items[0].sku if order.line_items else ""
            attrs = getattr(self._window, "sku_attr_map", {}).get(sku, {})
            pt = attrs.get("postage_type", "").lower()
            if "minilope" in pt:
                return "minilope"
            if "devilope" in pt:
                return "devilope"
            if "satchel" in pt:
                return "satchel"
            # Fall back to PrivateNotes keywords
            notes = (order.buyer_notes or "").lower()
            li_notes = (order.line_items[0].notes if order.line_items else "").lower()
            combined = notes + " " + li_notes
            if "mini." in combined:
                return "minilope"
            if "devil." in combined:
                return "devilope"
            if "satchel" in combined:
                return "satchel"
        return None

    def _show_order(self, idx: int):
        order = self._queue[idx]
        li = order.line_items[0]

        if isinstance(order, NetoOrder):
            name = (
                f"{order.ship_first_name} {order.ship_last_name}".strip()
                or order.customer_name
            )
            state = order.ship_state
        else:
            name = order.ship_name or order.buyer_name
            state = order.ship_state

        self._customer_label.configure(
            text=name + (f"  ({state})" if state else "")
        )

        platform = "Neto" if isinstance(order, NetoOrder) else "eBay"
        price = getattr(li, "unit_price", 0.0)
        self._order_info_label.configure(
            text=f"{platform}  •  Order {order.order_id}  •  ${price:.2f}"
        )

        self._sku_label.configure(text=f"SKU: {li.sku}")
        desc = getattr(li, "product_name", None) or getattr(li, "title", None) or "(no description)"
        self._desc_label.configure(text=desc)

        if isinstance(order, NetoOrder) and li.postage_type:
            self._existing_label.configure(
                text=f"Existing PostageType: {li.postage_type}"
            )
        elif isinstance(order, EbayOrder):
            attrs = getattr(self._window, "sku_attr_map", {}).get(li.sku, {})
            pt = attrs.get("postage_type", "")
            self._existing_label.configure(
                text=f"Neto product: PostageType={pt or '(none)'}" if attrs else ""
            )
        else:
            self._existing_label.configure(text="")

        # Reset image placeholder then load asynchronously.
        # Use placeholder (not None) to avoid GC-triggered TclError: image "pyimageN" doesn't exist
        self._displayed_img = None
        self._img_label.configure(image=self._placeholder_img, text="")
        self._img_generation = getattr(self, "_img_generation", 0) + 1
        gen = self._img_generation

        if isinstance(order, NetoOrder):
            sku = li.sku
            def _fetch_neto(s=sku, g=gen):
                try:
                    log.debug("Fetching image URL for SKU %s (gen %d)", s, g)
                    url_map = self._window.neto_client.get_product_images([s])
                    url = url_map.get(s, "")
                    log.debug("Image URL for %s: %r", s, url or "(none)")
                    if url:
                        self._load_image_async(url, g)
                except Exception as exc:
                    log.debug("Image URL fetch failed for %s: %s", s, exc)
            threading.Thread(target=_fetch_neto, daemon=True).start()
        else:
            # eBay order — try inline URL first, then fall back to get_item_images()
            item_id = getattr(li, "legacy_item_id", "")
            inline_url = li.image_url
            log.debug("eBay image: inline_url=%r  legacy_item_id=%r", inline_url, item_id)
            if inline_url:
                self._load_image_async(inline_url, gen)
            elif item_id and self._window.ebay_client:
                def _fetch_ebay(iid=item_id, g=gen):
                    try:
                        url_map = self._window.ebay_client.get_item_images([iid])
                        url = url_map.get(iid, "")
                        log.debug("eBay get_item_images %s → %r", iid, url or "(none)")
                        if url:
                            self._load_image_async(url, g)
                    except Exception as exc:
                        log.debug("eBay image fetch failed for %s: %s", iid, exc)
                threading.Thread(target=_fetch_ebay, daemon=True).start()
            else:
                log.debug("eBay item has no image_url and no legacy_item_id — skipping")

    def _load_image_async(self, url: str, generation: int):
        def _load():
            try:
                import requests as req
                from PIL import Image
                from io import BytesIO
                log.debug("Downloading image gen=%d  url=%s", generation, url)
                resp = req.get(url, timeout=5)
                resp.raise_for_status()
                img = Image.open(BytesIO(resp.content)).convert("RGB")
                img.thumbnail((200, 200))
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(200, 200))
                def _apply(i=ctk_img, g=generation):
                    current_gen = getattr(self, "_img_generation", 0)
                    log.debug("_apply called gen=%d current=%d match=%s", g, current_gen, g == current_gen)
                    if current_gen == g:
                        self._displayed_img = i  # Hold ref to prevent GC
                        self._img_label.configure(image=i, text="")
                self.after(0, _apply)
            except Exception as exc:
                log.debug("Image download/display failed gen=%d: %s", generation, exc)

        threading.Thread(target=_load, daemon=True).start()

    def _update_progress(self):
        remaining = len(self._queue) - self._current_idx
        self._progress_label.configure(
            text=f"{remaining} of {len(self._queue)} remaining"
        )

    def _on_classify(self, label: str):
        order = self._queue[self._current_idx]
        self._classifications[order.order_id] = label
        self._window.envelope_classifications[order.order_id] = label

        # Write-back to Neto product catalogue for both Neto and eBay orders
        sku = order.line_items[0].sku if order.line_items else ""
        dry_run = self._window.config.app.dry_run
        if sku:
            if label == "books":
                threading.Thread(
                    target=self._window.neto_client.update_item_shipping_category,
                    args=(sku, "4"),
                    kwargs={"dry_run": dry_run},
                    daemon=True,
                ).start()
            else:
                postage_type = _POSTAGE_TYPE_MAP.get(label, label.capitalize())
                threading.Thread(
                    target=self._window.neto_client.update_item_postage_type,
                    args=(sku, postage_type),
                    kwargs={"dry_run": dry_run},
                    daemon=True,
                ).start()

        # Books orders are handled entirely separately — remove from all subsequent steps
        if label == "books":
            oid = order.order_id
            self._window.envelope_classifications.pop(oid, None)
            self._window.neto_orders = [o for o in self._window.neto_orders if o.order_id != oid]
            self._window.ebay_orders = [o for o in self._window.ebay_orders if o.order_id != oid]

        self._advance()

    def _on_skip(self):
        self._advance()

    def _advance(self):
        self._current_idx += 1
        if self._current_idx >= len(self._queue):
            classified = len(self._classifications)
            self._status_label.configure(
                text=(
                    f"Classification complete — "
                    f"{classified} order{'s' if classified != 1 else ''} classified."
                )
            )
            self._status_frame.tkraise()
            self._next_btn.configure(state="normal")
        else:
            self._show_order(self._current_idx)
            self._update_progress()

    def _proceed(self):
        self._on_complete()


# ── Step 4: Envelope PDF Generation ──────────────────────────────────────────

class EnvelopePDFView(ctk.CTkFrame):
    """
    Step 4 — Generate Minilope and Devilope address PDFs and show
    open / print buttons.

    generate_pdfs() is called when this step becomes active.  It batches
    any pending Neto PostageType updates and then generates the PDFs in a
    background thread.
    """

    def __init__(self, master, window, on_complete, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_complete = on_complete
        self._on_back = on_back
        self._pdf_paths: dict[str, str | None] = {"minilope": None, "devilope": None}
        self._output_dir: str = r"\\SERVER\Project Folder\Order-Fulfillment-App\Envelopes"
        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        ctk.CTkLabel(
            self,
            text="Envelope PDFs",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(16, 12))

        center = ctk.CTkFrame(self, fg_color="transparent")
        center.pack(expand=True, fill="both", padx=20)

        self._status_label = ctk.CTkLabel(
            center,
            text="Generating PDFs…",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"),
        )
        self._status_label.pack(expand=True)

        # PDF result rows (hidden until generated)
        self._results_frame = ctk.CTkFrame(center, fg_color="transparent")

        for label, display in (("minilope", "Minilopes"), ("devilope", "Devilopes")):
            row = ctk.CTkFrame(self._results_frame, fg_color="transparent")
            row.pack(fill="x", pady=8)

            count_label = ctk.CTkLabel(
                row, text=f"{display}: —",
                font=ctk.CTkFont(size=13), width=160, anchor="w",
            )
            count_label.pack(side="left")

            open_btn = ctk.CTkButton(
                row, text=f"Open {display}",
                width=140, height=32, state="disabled",
                command=lambda lbl=label: self._open_pdf(lbl),
            )
            open_btn.pack(side="left", padx=(8, 4))

            print_btn = ctk.CTkButton(
                row, text=f"Print {display}",
                width=140, height=32, state="disabled",
                fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
                command=lambda lbl=label: self._print_pdf(lbl),
            )
            print_btn.pack(side="left", padx=(0, 0))

            # Store refs for later updates
            setattr(self, f"_count_{label}", count_label)
            setattr(self, f"_open_{label}", open_btn)
            setattr(self, f"_print_{label}", print_btn)

        # Open folder button
        folder_row = ctk.CTkFrame(self._results_frame, fg_color="transparent")
        folder_row.pack(fill="x", pady=(4, 0))
        self._open_folder_btn = ctk.CTkButton(
            folder_row, text="Open Envelopes Folder",
            width=200, height=32,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=self._open_folder,
        )
        self._open_folder_btn.pack(side="left")

        # Bottom nav
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", side="bottom", padx=20, pady=(8, 16))

        self._back_btn = ctk.CTkButton(
            bottom, text="← Back",
            width=100, fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            state="disabled",
            command=self._on_back,
        )
        self._back_btn.pack(side="left")

        self._next_btn = ctk.CTkButton(
            bottom, text="Next: Pick Zones →",
            width=200, state="disabled",
            command=self._on_complete,
        )
        self._next_btn.pack(side="right")

    # ── Entry point ───────────────────────────────────────────────────────

    def generate_pdfs(self):
        """Called by DailyOpsWindow each time this step is shown."""
        self._status_label.configure(
            text="Generating PDFs…"
        )
        self._results_frame.pack_forget()
        self._back_btn.configure(state="disabled")
        self._next_btn.configure(state="disabled")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        error = None
        try:
            self._pdf_paths = self._run_pdf_generation()
        except Exception as e:
            error = str(e)
        self.after(0, lambda: self._on_worker_done(error))

    def _run_pdf_generation(self) -> dict[str, str | None]:
        from src.envelope_pdf import generate_envelope_pdfs

        all_orders = self._window.neto_orders + self._window.ebay_orders
        classifications = {
            oid: label
            for oid, label in self._window.envelope_classifications.items()
            if label in ("minilope", "devilope")
        }
        output_dir = r"\\SERVER\Project Folder\Order-Fulfillment-App\Envelopes"
        self._output_dir = output_dir
        return generate_envelope_pdfs(all_orders, classifications, output_dir)

    def _on_worker_done(self, error: str | None):
        self._back_btn.configure(state="normal")
        self._next_btn.configure(state="normal")

        if error:
            self._status_label.configure(
                text=f"Error generating PDFs: {error}",
            )
            return

        self._status_label.configure(text="")
        self._results_frame.pack(expand=True, fill="both")

        # Update count labels and enable buttons
        classifications = self._window.envelope_classifications
        mini_count = sum(1 for v in classifications.values() if v == "minilope")
        devil_count = sum(1 for v in classifications.values() if v == "devilope")

        for label, count in (("minilope", mini_count), ("devilope", devil_count)):
            display = "Minilopes" if label == "minilope" else "Devilopes"
            getattr(self, f"_count_{label}").configure(
                text=f"{display}: {count} order{'s' if count != 1 else ''}"
            )
            path = self._pdf_paths.get(label)
            if path:
                getattr(self, f"_open_{label}").configure(state="normal")
                getattr(self, f"_print_{label}").configure(state="normal")
            else:
                getattr(self, f"_open_{label}").configure(state="disabled")
                getattr(self, f"_print_{label}").configure(state="disabled")

    # ── PDF actions ───────────────────────────────────────────────────────

    def _open_folder(self):
        folder = getattr(self, "_output_dir", None) or r"\\SERVER\Project Folder\Order-Fulfillment-App\Envelopes"
        import subprocess
        subprocess.Popen(["explorer", folder])

    def _open_pdf(self, label: str):
        path = self._pdf_paths.get(label)
        if path and os.path.exists(path):
            import subprocess
            subprocess.Popen(["start", "", path], shell=True)

    def _print_pdf(self, label: str):
        path = self._pdf_paths.get(label)
        if not path or not os.path.exists(path):
            return
        import subprocess
        import shutil

        # Try SumatraPDF — supports direct print with monochrome + actual-size
        # (1x = 100% / no scaling, so A5 PDF prints on A5 paper; mono = B&W)
        sumatra = _find_sumatra()
        if sumatra:
            subprocess.Popen([
                sumatra,
                "-print-to-default",
                "-print-settings", "1x,mono",
                "-silent",
                path,
            ])
        else:
            # Fallback: open print dialog via default viewer
            log.warning("SumatraPDF not found — falling back to shell print dialog")
            subprocess.Popen(["start", "/print", path], shell=True)
