from __future__ import annotations

import io
import threading
import tkinter as tk
from datetime import date
from tkinter import messagebox

import customtkinter as ctk
import requests
from PIL import Image, ImageTk

from src.order_collator import CollatedGroup, _normalize_street1


_PLACEHOLDER_SIZE = (50, 50)

_SHIPPING_METHODS = [
    "Allied Express",
    "Aramex",
    "Australia Post",
    "Bonds Couriers",
    "Courier's Please",
    "DAI Post",
    "Toll",
]


class CollatedDetailView(ctk.CTkFrame):
    """
    Full-screen overlay showing all sub-orders in a CollatedGroup.

    Layout:
        Header bar  (← Back | title | Ungroup)
        Scrollable content:
            Shared address block
            Per-order card × N  (line items + notes + pricing + tracking + actions)
        Footer  (Book Freight for All | Mark All as Sent)
    """

    def __init__(
        self,
        master,
        *,
        group: CollatedGroup,
        window,
        dry_run: bool = True,
        on_back,
        on_ungroup,
        on_book_freight,
    ):
        super().__init__(master, fg_color="transparent")
        self._group = group
        self._window = window
        self._dry_run = dry_run
        self._on_back = on_back
        self._on_ungroup = on_ungroup          # on_ungroup(order_ids: list[str])
        self._on_book_freight = on_book_freight  # on_book_freight(order_id, platform, for_all)

        # One dict per sub-order (populated by _build_order_card)
        self._cards: list[dict] = []

        # Image resources
        self._full_images: dict[str, bytes] = {}
        self._image_refs: list = []

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_ui()

    # ── Top-level layout ───────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_scroll_area()
        self._build_footer()

    # ── Header ─────────────────────────────────────────────────────────────────

    def _build_header(self):
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray85", "gray20"), height=50)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            header, text="← Back", width=90,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=self._on_back,
        ).grid(row=0, column=0, padx=(12, 8), pady=10)

        first = self._group.orders[0]
        is_neto = hasattr(first, "date_placed")
        if is_neto:
            cust = f"{first.ship_first_name} {first.ship_last_name}".strip() or first.customer_name
        else:
            cust = first.ship_name or first.buyer_name

        n = len(self._group.orders)
        ctk.CTkLabel(
            header,
            text=f"Collated Order  —  {n} orders for {cust}",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=1, padx=8, pady=10, sticky="w")

        ctk.CTkButton(
            header, text="Ungroup", width=90,
            fg_color=("gray60", "gray35"), hover_color=("gray50", "gray30"),
            command=self._confirm_ungroup,
        ).grid(row=0, column=2, padx=(8, 12), pady=10)

    # ── Scroll area ────────────────────────────────────────────────────────────

    def _build_scroll_area(self):
        self._scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._scroll.grid(row=1, column=0, sticky="nsew", padx=16, pady=(8, 4))
        self._scroll.grid_columnconfigure(0, weight=1)

        self._build_shared_address(self._scroll)

        for i, order in enumerate(self._group.orders):
            card = self._build_order_card(self._scroll, order, i)
            self._cards.append(card)

        # Start background image fetches now that all label widgets exist
        self._start_image_fetches()

    # ── Shared address ─────────────────────────────────────────────────────────

    def _build_shared_address(self, parent):
        first = self._group.orders[0]
        is_neto = hasattr(first, "date_placed")

        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"),
                             corner_radius=6)
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        ctk.CTkLabel(frame, text="Shared Shipping Address",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4))

        if is_neto:
            name = f"{first.ship_first_name} {first.ship_last_name}".strip()
            lines = [
                name,
                first.ship_company,
                first.ship_street1,
                first.ship_street2,
                f"{first.ship_city} {first.ship_state} {first.ship_postcode}".strip(),
                first.ship_country,
                first.ship_phone,
            ]
        else:
            # Strip eBay-injected tracking prefix before displaying
            street1 = _normalize_street1(first.ship_street1)
            lines = [
                first.ship_name,
                street1,
                first.ship_street2,
                f"{first.ship_city} {first.ship_state} {first.ship_postcode}".strip(),
                first.ship_country,
                first.ship_phone,
            ]

        for line in lines:
            if not line:
                continue
            row = ctk.CTkFrame(frame, fg_color="transparent")
            row.pack(fill="x", padx=10, pady=1)
            ctk.CTkLabel(row, text=line, font=ctk.CTkFont(size=13), anchor="w").pack(
                side="left", fill="x", expand=True)
            btn = ctk.CTkButton(row, text="Copy", width=45, height=22,
                                font=ctk.CTkFont(size=10),
                                fg_color="gray50", hover_color="gray40")
            btn.configure(command=lambda t=line, b=btn: self._copy_to_clipboard(t, b))
            btn.pack(side="right", padx=4)

        all_text = "\n".join(l for l in lines if l)
        copy_all_btn = ctk.CTkButton(frame, text="Copy All", width=70, height=26,
                                     font=ctk.CTkFont(size=11))
        copy_all_btn.configure(
            command=lambda b=copy_all_btn: self._copy_to_clipboard(all_text, b, "Copy All"))
        copy_all_btn.pack(anchor="e", padx=10, pady=(4, 8))

    # ── Per-order card ─────────────────────────────────────────────────────────

    def _build_order_card(self, parent, order, idx: int) -> dict:
        is_neto = hasattr(order, "date_placed")
        platform = order.sales_channel if is_neto else "eBay"

        card_frame = ctk.CTkFrame(parent, corner_radius=8, border_width=1,
                                  border_color=("gray70", "gray35"))
        card_frame.grid(row=idx + 1, column=0, sticky="ew", pady=(0, 12))
        card_frame.grid_columnconfigure(0, weight=1)

        # ── Title bar ────────────────────────────────────────────────────────
        title_bar = ctk.CTkFrame(card_frame, fg_color=("gray80", "gray25"), corner_radius=0)
        title_bar.grid(row=0, column=0, sticky="ew")
        title_inner = ctk.CTkFrame(title_bar, fg_color="transparent")
        title_inner.pack(side="left", padx=12, pady=6)

        ctk.CTkLabel(title_inner, text=order.order_id,
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=(0, 4))

        _copy_id_btn = ctk.CTkButton(
            title_inner, text="Copy", width=45, height=22, font=ctk.CTkFont(size=10),
            fg_color="gray50", hover_color="gray40")
        _copy_id_btn.configure(
            command=lambda oid=order.order_id, b=_copy_id_btn:
            self._copy_to_clipboard(oid, b))
        _copy_id_btn.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(
            title_inner, text=platform,
            font=ctk.CTkFont(size=11),
            fg_color=("dodgerblue3", "dodgerblue4"),
            corner_radius=4, text_color="white",
        ).pack(side="left", ipadx=5, ipady=1)

        if is_neto:
            d = order.date_paid or order.date_placed
            date_str = d.strftime("%d/%m/%Y") if d else ""
        else:
            date_str = order.creation_date.strftime("%d/%m/%Y") if order.creation_date else ""
        if date_str:
            ctk.CTkLabel(title_inner, text=date_str, font=ctk.CTkFont(size=11),
                         text_color="gray50").pack(side="left", padx=(12, 0))

        # ── Line items ───────────────────────────────────────────────────────
        li_section = ctk.CTkFrame(card_frame, fg_color="transparent")
        li_section.grid(row=1, column=0, sticky="ew", padx=12, pady=(10, 0))

        ctk.CTkLabel(li_section, text="Line Items",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0, 4))

        # Column headers
        hdr = ctk.CTkFrame(li_section, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 2))
        for text, w in [("", 54), ("SKU", 130), ("Description", 240), ("Qty", 40), ("Price", 70)]:
            ctk.CTkLabel(hdr, text=text, width=w, font=ctk.CTkFont(size=11, weight="bold"),
                         anchor="w").pack(side="left", padx=(0, 6))

        items_frame = ctk.CTkFrame(li_section, fg_color="transparent")
        items_frame.pack(fill="x", pady=(0, 8))

        neto_img_pending: dict = {}
        ebay_img_pending: dict = {}
        ebay_note_widgets: list = []

        if is_neto:
            for li in order.line_items:
                img_lbl = self._build_item_row(
                    items_frame, li.sku, li.product_name, li.quantity, li.unit_price,
                    order, li, is_neto)
                if li.sku:
                    neto_img_pending[li.sku] = img_lbl
        else:
            for li in order.line_items:
                img_lbl = self._build_item_row(
                    items_frame, li.sku, li.title, li.quantity, li.unit_price,
                    order, li, is_neto)
                if li.legacy_item_id:
                    ebay_img_pending[li.legacy_item_id] = img_lbl

                # Inline per-item note editor (eBay private notes)
                note_row = ctk.CTkFrame(items_frame, fg_color=("gray90", "gray25"),
                                        corner_radius=4)
                note_row.pack(fill="x", padx=(54, 0), pady=(0, 6))
                ctk.CTkLabel(note_row, text="Notes:", width=50,
                             font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(
                    side="left", padx=(6, 4), pady=4)
                entry = ctk.CTkEntry(note_row, font=ctk.CTkFont(size=11), height=26,
                                     text_color="#f5c518")
                entry.pack(side="left", fill="x", expand=True, pady=4)
                if li.notes:
                    entry.insert(0, li.notes)
                char_lbl = ctk.CTkLabel(note_row,
                                        text=f"{len(li.notes or '')}/255",
                                        width=50, font=ctk.CTkFont(size=10),
                                        text_color="gray50")
                char_lbl.pack(side="left", padx=(4, 0), pady=4)
                entry.bind("<KeyRelease>",
                           lambda e, en=entry, cl=char_lbl:
                           self._update_char_count(en, cl))
                save_btn = ctk.CTkButton(note_row, text="Save", width=55, height=24,
                                         font=ctk.CTkFont(size=11))
                save_btn.configure(
                    command=lambda l=li, o=order, en=entry, b=save_btn:
                    self._save_ebay_note(l, o, en, b))
                save_btn.pack(side="left", padx=(4, 6), pady=4)
                ebay_note_widgets.append((li, entry, save_btn))

        # ── Notes section ────────────────────────────────────────────────────
        notes_frame = ctk.CTkFrame(card_frame, fg_color="transparent")
        notes_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))

        # Build preliminary card dict so note helpers can store widget refs
        card: dict = {
            "order": order,
            "platform": platform,
            "is_neto": is_neto,
            "tracking_entry": None,
            "carrier_combo": None,
            "send_btn": None,
            "status_lbl": None,
            "neto_img_pending": neto_img_pending,
            "ebay_img_pending": ebay_img_pending,
            "ebay_note_widgets": ebay_note_widgets,
        }

        self._build_notes_section(notes_frame, card)

        # ── Pricing ──────────────────────────────────────────────────────────
        if is_neto:
            total = getattr(order, "grand_total", 0.0) or 0.0
            shipping_cost = getattr(order, "shipping_total", 0.0) or 0.0
        else:
            total = getattr(order, "order_total", 0.0) or 0.0
            shipping_cost = getattr(order, "shipping_cost", 0.0) or 0.0

        pricing_row = ctk.CTkFrame(card_frame, fg_color="transparent")
        pricing_row.grid(row=3, column=0, sticky="ew", padx=12, pady=(2, 4))
        ctk.CTkLabel(pricing_row, text=f"Shipping: ${shipping_cost:.2f}",
                     font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 20))
        ctk.CTkLabel(pricing_row, text=f"Order Total: ${total:.2f}",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")

        # ── Tracking / carrier row ───────────────────────────────────────────
        track_row = ctk.CTkFrame(card_frame, fg_color="transparent")
        track_row.grid(row=4, column=0, sticky="ew", padx=12, pady=(4, 4))

        ctk.CTkLabel(track_row, text="Tracking:", font=ctk.CTkFont(size=12)).pack(side="left")
        tracking_entry = ctk.CTkEntry(track_row, width=160, font=ctk.CTkFont(size=12))
        tracking_entry.pack(side="left", padx=(4, 12))
        ctk.CTkLabel(track_row, text="Carrier:", font=ctk.CTkFont(size=12)).pack(side="left")
        carrier_combo = ctk.CTkComboBox(
            track_row, values=_SHIPPING_METHODS, width=160, font=ctk.CTkFont(size=12))
        carrier_combo.set("")
        carrier_combo.pack(side="left", padx=(4, 0))

        card["tracking_entry"] = tracking_entry
        card["carrier_combo"] = carrier_combo

        # ── Action buttons row ───────────────────────────────────────────────
        btn_row = ctk.CTkFrame(card_frame, fg_color="transparent")
        btn_row.grid(row=5, column=0, sticky="ew", padx=12, pady=(2, 12))

        status_lbl = ctk.CTkLabel(btn_row, text="", font=ctk.CTkFont(size=12))
        status_lbl.pack(side="left")

        send_btn = ctk.CTkButton(
            btn_row, text="Mark as Sent", width=120,
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
        )
        send_btn.pack(side="right", padx=(6, 0))

        book_btn = ctk.CTkButton(
            btn_row, text="Book Freight", width=120,
            fg_color=("dodgerblue3", "dodgerblue4"),
        )
        book_btn.pack(side="right")

        card["send_btn"] = send_btn
        card["status_lbl"] = status_lbl

        send_btn.configure(command=lambda c=card: self._mark_card_as_sent(c))
        book_btn.configure(
            command=lambda oid=order.order_id, plat=platform:
            self._on_book_freight(oid, plat, False))

        return card

    # ── Item row builder ───────────────────────────────────────────────────────

    def _build_item_row(self, parent, sku, desc, qty, price, order, line_item,
                        is_neto) -> ctk.CTkLabel:
        """Build a single line-item row. Returns the image placeholder label."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=1)

        img_label = ctk.CTkLabel(row, text="·", width=50, height=50,
                                 text_color="gray50", font=ctk.CTkFont(size=20))
        img_label.pack(side="left", padx=(0, 4))

        price_str = f"${price:.2f}" if price else ""
        ctk.CTkLabel(row, text=sku or "—", width=130, anchor="w",
                     wraplength=130).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(row, text=desc or "—", width=240, anchor="w",
                     wraplength=240).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(row, text=str(qty), width=40, anchor="w").pack(side="left", padx=(0, 6))
        ctk.CTkLabel(row, text=price_str, width=70, anchor="w").pack(side="left")

        # Action sub-row: Left in Waiting Area + On PO
        action_row = ctk.CTkFrame(parent, fg_color="transparent")
        action_row.pack(fill="x", padx=(54, 0), pady=(0, 4))

        row_status = ctk.CTkLabel(action_row, text="", font=ctk.CTkFont(size=11),
                                  text_color=("gray40", "gray60"))

        wait_btn = ctk.CTkButton(
            action_row, text="Left in Waiting Area", width=150, height=26,
            font=ctk.CTkFont(size=11),
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
        )
        wait_btn.pack(side="left", padx=(0, 6))
        row_status.pack(side="left", padx=(0, 6))
        wait_btn.configure(
            command=lambda li=line_item, o=order, b=wait_btn, s=row_status:
            self._on_waiting_area_clicked(li, o, b, s))

        musipos = getattr(self._window, "musipos_client", None)
        if musipos is not None:
            po_btn = ctk.CTkButton(
                action_row, text="On PO", width=70, height=26,
                font=ctk.CTkFont(size=11),
                fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            )
            po_btn.configure(
                command=lambda li=line_item, o=order, b=po_btn, s=row_status:
                self._on_add_to_po_clicked(li, o, b, s))
            po_btn.pack(side="left")

        return img_label

    # ── Notes section ──────────────────────────────────────────────────────────

    def _build_notes_section(self, parent, card: dict):
        order = card["order"]
        is_neto = card["is_neto"]

        if is_neto:
            ctk.CTkLabel(parent, text="Notes",
                         font=ctk.CTkFont(size=12, weight="bold")).pack(
                anchor="w", pady=(4, 2))
            card["_notes_content_parent"] = parent
            content = ctk.CTkFrame(parent, fg_color="transparent")
            content.pack(fill="x")
            card["_notes_content_frame"] = content
            self._populate_neto_notes(content, card)
        else:
            buyer_notes = getattr(order, "buyer_notes", "") or ""
            if buyer_notes:
                ctk.CTkLabel(parent, text=f"Buyer notes: {buyer_notes}",
                             font=ctk.CTkFont(size=12), anchor="w", wraplength=700,
                             text_color="#f5c518").pack(fill="x", pady=(2, 4))

    def _populate_neto_notes(self, parent, card: dict):
        """Render Neto notes content (sticky notes + add-note box) into *parent*."""
        order = card["order"]

        if order.delivery_instruction:
            ctk.CTkLabel(parent,
                         text=f"Delivery: {order.delivery_instruction}",
                         font=ctk.CTkFont(size=12), anchor="w", wraplength=700,
                         text_color="#f5c518").pack(fill="x", pady=2)

        if order.sticky_notes:
            ctk.CTkLabel(parent, text="Sticky Notes:",
                         font=ctk.CTkFont(size=12, weight="bold")).pack(
                anchor="w", pady=(4, 2))

            def _key(n):
                try:
                    return int(n.get("StickyNoteID") or 0)
                except (ValueError, TypeError):
                    return 0

            for note in sorted(order.sticky_notes, key=_key, reverse=True):
                title = note.get("Title", "")
                desc = note.get("Description", "")
                date_str = note.get("DateAdded") or note.get("DateCreated") or ""
                header = f"[{date_str}] " if date_str else ""
                text = f"{header}{title}: {desc}" if title else f"{header}{desc}"
                tb = ctk.CTkTextbox(
                    parent, height=50, font=ctk.CTkFont(size=12),
                    text_color="#f5c518", fg_color=("gray90", "gray25"), border_width=0)
                tb.insert("1.0", text)
                tb.configure(state="disabled")
                tb.pack(fill="x", pady=1)

        if order.internal_notes:
            ctk.CTkLabel(parent, text=f"Internal: {order.internal_notes}",
                         font=ctk.CTkFont(size=12), anchor="w", wraplength=700,
                         text_color="#f5c518").pack(fill="x", pady=2)

        ctk.CTkLabel(parent, text="Add Sticky Note:",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(6, 2))
        note_tb = ctk.CTkTextbox(parent, height=56, font=ctk.CTkFont(size=12))
        note_tb.pack(fill="x", pady=(0, 4))
        card["_note_textbox"] = note_tb

        add_btn = ctk.CTkButton(parent, text="Add Note", width=90, height=28,
                                command=lambda c=card: self._add_neto_note(c))
        add_btn.pack(anchor="e", pady=(0, 4))
        card["_add_note_btn"] = add_btn

    def _rebuild_neto_notes(self, card: dict):
        """Destroy + recreate the dynamic notes content for a Neto card."""
        old = card.get("_notes_content_frame")
        if old:
            old.destroy()
        parent = card["_notes_content_parent"]
        content = ctk.CTkFrame(parent, fg_color="transparent")
        content.pack(fill="x")
        card["_notes_content_frame"] = content
        self._populate_neto_notes(content, card)

    # ── Footer ─────────────────────────────────────────────────────────────────

    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), corner_radius=0, height=56)
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)
        footer.grid_columnconfigure(0, weight=1)

        self._all_status_lbl = ctk.CTkLabel(footer, text="", font=ctk.CTkFont(size=12))
        self._all_status_lbl.place(relx=0.02, rely=0.5, anchor="w")

        btn_inner = ctk.CTkFrame(footer, fg_color="transparent")
        btn_inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkButton(
            btn_inner, text="Book Freight for All", width=180,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._book_freight_for_all,
        ).pack(side="left", padx=(0, 12))

        self._mark_all_btn = ctk.CTkButton(
            btn_inner, text="Mark All as Sent", width=160,
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
            command=self._confirm_mark_all,
        )
        self._mark_all_btn.pack(side="left")

    # ── Image fetching ─────────────────────────────────────────────────────────

    def _start_image_fetches(self):
        neto_client = self._window.neto_client
        ebay_client = self._window.ebay_client
        for card in self._cards:
            if card["is_neto"] and card["neto_img_pending"] and neto_client:
                pending = dict(card["neto_img_pending"])
                threading.Thread(
                    target=self._fetch_neto_images, args=(pending, neto_client),
                    daemon=True).start()
            elif not card["is_neto"] and card["ebay_img_pending"] and ebay_client:
                pending = dict(card["ebay_img_pending"])
                threading.Thread(
                    target=self._fetch_ebay_images, args=(pending, ebay_client),
                    daemon=True).start()

    def _fetch_neto_images(self, pending: dict, client):
        try:
            url_map = client.get_product_images(list(pending))
            for sku, url in url_map.items():
                label = pending.get(sku)
                if url and label:
                    self._download_and_show_image(url, label)
        except Exception as exc:
            print(f"[CollatedDetail] Neto image fetch failed: {exc}")

    def _fetch_ebay_images(self, pending: dict, client):
        try:
            url_map = client.get_item_images(list(pending))
            for item_id, url in url_map.items():
                label = pending.get(item_id)
                if url and label:
                    self._download_and_show_image(url, label)
        except Exception as exc:
            print(f"[CollatedDetail] eBay image fetch failed: {exc}")

    def _download_and_show_image(self, url: str, label: ctk.CTkLabel):
        try:
            resp = requests.get(url, timeout=10)
            resp.raise_for_status()
            data = resp.content
            self._full_images[url] = data
            img = Image.open(io.BytesIO(data))
            img.thumbnail(_PLACEHOLDER_SIZE, Image.Resampling.LANCZOS)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=_PLACEHOLDER_SIZE)
            self._image_refs.append(ctk_img)

            def _update(u=url, lb=label, ci=ctk_img):
                lb.configure(image=ci, text="")
                lb.bind("<Button-1>", lambda e, uu=u: self._open_image_large(uu))
                lb.configure(cursor="hand2")

            label.after(0, _update)
        except Exception as exc:
            print(f"[CollatedDetail] Image download failed {url!r}: {exc}")

    def _open_image_large(self, url: str):
        data = self._full_images.get(url)
        if not data:
            return
        try:
            img = Image.open(io.BytesIO(data))
            img.thumbnail((600, 600), Image.Resampling.LANCZOS)
            top = tk.Toplevel(self.winfo_toplevel())
            top.title("Image Preview")
            top.resizable(False, False)
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(top, image=photo, cursor="hand2")
            lbl.image = photo
            lbl.pack()
            tk.Label(top, text="Click to close", fg="gray60").pack(pady=(0, 4))
            lbl.bind("<Button-1>", lambda e: top.destroy())
            top.bind("<Escape>", lambda e: top.destroy())
            top.update_idletasks()
            w, h = top.winfo_width(), top.winfo_height()
            sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
            top.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")
        except Exception:
            pass

    # ── Per-item action handlers ───────────────────────────────────────────────

    def _on_waiting_area_clicked(self, line_item, order, btn, status_lbl):
        btn.configure(state="disabled")
        is_neto = hasattr(order, "date_placed")
        note_text = f"{line_item.sku} left in waiting area"
        status_lbl.configure(text="Adding note…", text_color=("gray40", "gray60"))

        def _work():
            try:
                if is_neto and self._window.neto_client:
                    dated = f"[{date.today().strftime('%d/%m/%Y')}] {note_text}"
                    self._window.neto_client.add_sticky_note(
                        order.order_id, title="Item Status", description=dated,
                        dry_run=self._dry_run)
                    self._refresh_neto_notes_for(order)
                elif not is_neto and self._window.ebay_client:
                    self._window.ebay_client.set_private_notes(
                        item_id=line_item.legacy_item_id,
                        transaction_id=line_item.legacy_transaction_id,
                        note_text=note_text[:255],
                        dry_run=self._dry_run,
                    )
                    line_item.notes = note_text[:255]
                suffix = " [DRY RUN]" if self._dry_run else ""
                self.after(0, lambda: status_lbl.configure(
                    text=f"✓ Note added{suffix}", text_color="green"))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: (
                    status_lbl.configure(text=f"Error: {err}", text_color="red"),
                    btn.configure(state="normal"),
                ))

        threading.Thread(target=_work, daemon=True).start()

    def _on_add_to_po_clicked(self, line_item, order, btn, status_lbl):
        btn.configure(state="disabled")
        from src.gui.musipos_po_dialog import MusiposPODialog
        musipos = getattr(self._window, "musipos_client", None)
        if musipos is None:
            btn.configure(state="normal")
            return

        def _on_success(po_result):
            self._on_po_added(line_item, order, po_result, btn, status_lbl)

        def _on_note_only():
            self._on_po_note_only(line_item, order, btn, status_lbl)

        def _on_close():
            if btn.cget("state") == "disabled":
                btn.configure(state="normal")

        qty = getattr(line_item, "quantity", 1) or 1
        dialog = MusiposPODialog(
            self.winfo_toplevel(),
            neto_sku=line_item.sku,
            product_name=(getattr(line_item, "product_name", None)
                          or getattr(line_item, "title", "") or ""),
            order_qty=qty,
            musipos_client=musipos,
            suppliers_config=self._window.config.suppliers,
            dry_run=self._dry_run,
            on_success=_on_success,
            on_note_only=_on_note_only,
        )
        dialog.protocol("WM_DELETE_WINDOW", lambda: (dialog.destroy(), _on_close()))

    def _on_po_added(self, line_item, order, po_result, btn, status_lbl):
        note_text = f"{line_item.sku} on PO"
        self._add_order_note(note_text, order, line_item)
        suffix = " [DRY RUN]" if self._dry_run else ""
        status_lbl.configure(text=f"✓ Added to PO #{po_result['po_no']}{suffix}",
                              text_color="green")
        btn.configure(text="On PO ✓")

    def _on_po_note_only(self, line_item, order, btn, status_lbl):
        note_text = f"{line_item.sku} on PO"
        self._add_order_note(note_text, order, line_item)
        status_lbl.configure(text="✓ Note added", text_color="green")
        btn.configure(state="normal")

    def _add_order_note(self, note_text: str, order, line_item=None):
        """Fire-and-forget: add a note to the given order and refresh notes display."""
        is_neto = hasattr(order, "date_placed")

        def _work():
            try:
                if is_neto and self._window.neto_client:
                    dated = f"[{date.today().strftime('%d/%m/%Y')}] {note_text}"
                    self._window.neto_client.add_sticky_note(
                        order.order_id, title="Item Status", description=dated,
                        dry_run=self._dry_run)
                    self._refresh_neto_notes_for(order)
                elif not is_neto and self._window.ebay_client:
                    target = line_item or (order.line_items[0] if order.line_items else None)
                    if target:
                        self._window.ebay_client.set_private_notes(
                            item_id=target.legacy_item_id,
                            transaction_id=target.legacy_transaction_id,
                            note_text=note_text[:255],
                            dry_run=self._dry_run,
                        )
            except Exception as exc:
                print(f"[CollatedDetail] Failed to add note: {exc}")

        threading.Thread(target=_work, daemon=True).start()

    def _refresh_neto_notes_for(self, order):
        """Re-fetch order notes from Neto and rebuild the card's notes section.
        Safe to call from a background thread."""
        try:
            fresh = self._window.neto_client.get_orders_by_ids([order.order_id])
            if fresh:
                order.sticky_notes = fresh[0].sticky_notes
                order.delivery_instruction = fresh[0].delivery_instruction
                order.internal_notes = fresh[0].internal_notes
        except Exception:
            pass
        card = next((c for c in self._cards if c["order"].order_id == order.order_id), None)
        if card:
            self.after(0, lambda c=card: self._rebuild_neto_notes(c))

    def _add_neto_note(self, card: dict):
        """Save the text from the card's Add Sticky Note box."""
        tb = card.get("_note_textbox")
        if tb is None:
            return
        text = tb.get("1.0", "end").strip()
        if not text:
            return
        order = card["order"]
        dated = f"[{date.today().strftime('%d/%m/%Y')}] {text}"
        add_btn = card.get("_add_note_btn")
        if add_btn:
            add_btn.configure(state="disabled", text="Saving…")
        try:
            self._window.neto_client.add_sticky_note(
                order.order_id, title="Packing Note", description=dated,
                dry_run=self._dry_run)
            if self._dry_run:
                if add_btn:
                    add_btn.configure(state="disabled", text="Note Added")
                messagebox.showinfo("Dry Run",
                                    f"[DRY RUN] Sticky note would be added:\n{dated}",
                                    parent=self.winfo_toplevel())
            else:
                threading.Thread(
                    target=self._refresh_neto_notes_for, args=(order,),
                    daemon=True).start()
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to add note:\n{exc}",
                                 parent=self.winfo_toplevel())
            if add_btn:
                add_btn.configure(state="normal", text="Add Note")

    def _save_ebay_note(self, line_item, order, entry, btn):
        text = entry.get().strip()
        if not text:
            return
        if len(text) > 255:
            messagebox.showwarning(
                "Too Long",
                f"eBay notes are limited to 255 characters.\nCurrent length: {len(text)}",
                parent=self.winfo_toplevel())
            return
        if not line_item.legacy_item_id or not line_item.legacy_transaction_id:
            messagebox.showerror("Error", "Missing eBay item/transaction IDs.",
                                 parent=self.winfo_toplevel())
            return
        btn.configure(state="disabled", text="Saving…")
        try:
            self._window.ebay_client.set_private_notes(
                item_id=line_item.legacy_item_id,
                transaction_id=line_item.legacy_transaction_id,
                note_text=text,
                dry_run=self._dry_run,
            )
            line_item.notes = text
            if self._dry_run:
                btn.configure(state="disabled", text="Saved")
                self.after(2000, lambda b=btn: b.configure(state="normal", text="Save"))
                messagebox.showinfo("Dry Run", f"[DRY RUN] Note would be set:\n{text}",
                                    parent=self.winfo_toplevel())
            else:
                btn.configure(text="Refreshing…")

                def _fetch():
                    try:
                        fresh = self._window.ebay_client.get_orders_by_ids([order.order_id])
                        self.after(0, lambda: self._on_ebay_note_refreshed(fresh, order, btn))
                    except Exception:
                        self.after(0, lambda b=btn: b.configure(state="normal", text="Save"))

                threading.Thread(target=_fetch, daemon=True).start()
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to save note:\n{exc}",
                                 parent=self.winfo_toplevel())
            btn.configure(state="normal", text="Save")

    def _on_ebay_note_refreshed(self, fresh_orders: list, order, btn):
        if fresh_orders:
            fresh = fresh_orders[0]
            fresh_map = {
                (li.legacy_item_id, li.legacy_transaction_id): li
                for li in fresh.line_items
            }
            card = next((c for c in self._cards if c["order"].order_id == order.order_id), None)
            if card:
                for li, entry_widget, _b in card["ebay_note_widgets"]:
                    fli = fresh_map.get((li.legacy_item_id, li.legacy_transaction_id))
                    if fli:
                        li.notes = fli.notes
                        entry_widget.delete(0, "end")
                        entry_widget.insert(0, fli.notes or "")
        if btn:
            btn.configure(state="normal", text="Save")

    # ── Char-count helper ──────────────────────────────────────────────────────

    def _update_char_count(self, entry, label):
        count = len(entry.get().strip())
        label.configure(text=f"{count}/255",
                        text_color="red" if count > 255 else "gray50")

    # ── Per-card dispatch ──────────────────────────────────────────────────────

    def _mark_card_as_sent(self, card: dict):
        order = card["order"]
        tracking = card["tracking_entry"].get().strip()
        carrier = card["carrier_combo"].get().strip()

        if not tracking:
            if not messagebox.askyesno(
                "No Tracking",
                f"Send order {order.order_id} without a tracking number?",
                parent=self.winfo_toplevel(),
            ):
                return

        card["send_btn"].configure(state="disabled", text="Sending…")
        card["status_lbl"].configure(text="")

        def _work():
            ok, msg = self._dispatch_order(order, tracking, carrier)
            self.after(0, lambda: self._on_card_sent(card, ok, msg))

        threading.Thread(target=_work, daemon=True).start()

    def _dispatch_order(self, order, tracking: str, carrier: str) -> tuple:
        try:
            is_neto = hasattr(order, "date_placed")
            if is_neto and self._window.neto_client:
                self._window.neto_client.update_order_status(
                    order.order_id, new_status="Dispatched",
                    tracking_number=tracking, shipping_method=carrier,
                    line_item_skus=[li.sku for li in order.line_items],
                    dry_run=self._dry_run,
                )
            elif not is_neto and self._window.ebay_client:
                self._window.ebay_client.create_shipping_fulfillment(
                    order.order_id, line_items=order.line_items,
                    tracking_number=tracking, carrier=carrier,
                    dry_run=self._dry_run,
                )
            msg = "[DRY RUN] Marked as sent" if self._dry_run else "Sent"
            return True, msg
        except Exception as exc:
            return False, str(exc)

    def _on_card_sent(self, card: dict, ok: bool, msg: str):
        if ok:
            card["send_btn"].configure(text="SENT", fg_color="gray50")
            card["tracking_entry"].configure(state="disabled")
            card["carrier_combo"].configure(state="disabled")
            card["status_lbl"].configure(
                text=msg, text_color="orange" if self._dry_run else "green")
        else:
            card["send_btn"].configure(state="normal", text="Mark as Sent")
            card["status_lbl"].configure(text=f"Error: {msg}", text_color="red")

    # ── Tracking fill (called from ResultsView after freight booking) ──────────

    def set_tracking_for(self, order_id: str, tracking: str, carrier: str):
        """Fill tracking/carrier on the card for a specific sub-order."""
        for card in self._cards:
            if card["order"].order_id == order_id:
                _fill_tracking(card, tracking, carrier)
                return

    def set_tracking_all(self, tracking: str, carrier: str):
        """Fill tracking/carrier on every sub-order card."""
        for card in self._cards:
            _fill_tracking(card, tracking, carrier)

    # ── Bulk footer actions ────────────────────────────────────────────────────

    def _book_freight_for_all(self):
        first = self._group.orders[0]
        is_neto = hasattr(first, "date_placed")
        platform = first.sales_channel if is_neto else "eBay"
        self._on_book_freight(first.order_id, platform, True)

    def _confirm_mark_all(self):
        n = len(self._cards)
        if not messagebox.askyesno("Mark All as Sent",
                                   f"Mark all {n} orders as dispatched?",
                                   parent=self.winfo_toplevel()):
            return
        self._mark_all_as_sent()

    def _mark_all_as_sent(self):
        self._mark_all_btn.configure(state="disabled", text="Sending…")
        self._all_status_lbl.configure(text="")

        cards_to_send = [c for c in self._cards
                         if c["send_btn"].cget("state") != "disabled"]
        if not cards_to_send:
            self._mark_all_btn.configure(state="normal", text="Mark All as Sent")
            return

        def _work():
            errors = []
            for i, card in enumerate(cards_to_send):
                order = card["order"]
                tracking = card["tracking_entry"].get().strip()
                carrier = card["carrier_combo"].get().strip()
                ok, msg = self._dispatch_order(order, tracking, carrier)
                self.after(0, lambda c=card, o=ok, m=msg: self._on_card_sent(c, o, m))
                if not ok:
                    errors.append(f"{order.order_id}: {msg}")
                progress = f"{i + 1}/{len(cards_to_send)}"
                self.after(0, lambda p=progress: self._all_status_lbl.configure(text=p))
            self.after(0, lambda: self._on_all_sent(errors))

        threading.Thread(target=_work, daemon=True).start()

    def _on_all_sent(self, errors: list):
        self._mark_all_btn.configure(state="normal", text="Mark All as Sent")
        if errors:
            self._all_status_lbl.configure(text=f"{len(errors)} failed", text_color="red")
        else:
            colour = "orange" if self._dry_run else "green"
            msg = "[DRY RUN] All sent" if self._dry_run else "All sent"
            self._all_status_lbl.configure(text=msg, text_color=colour)

    # ── Ungroup ────────────────────────────────────────────────────────────────

    def _confirm_ungroup(self):
        n = len(self._group.orders)
        if not messagebox.askyesno("Ungroup Orders",
                                   f"Split this group back into {n} individual orders?",
                                   parent=self.winfo_toplevel()):
            return
        self._on_ungroup(self._group.order_ids)
        self._on_back()

    # ── Clipboard ─────────────────────────────────────────────────────────────

    def _copy_to_clipboard(self, text: str, btn=None, original_text: str = "Copy"):
        self.clipboard_clear()
        self.clipboard_append(text)
        if btn is not None:
            btn.configure(text="Copied!")
            self.after(2000, lambda: btn.configure(text=original_text))


# ── Module-level helper ────────────────────────────────────────────────────────

def _fill_tracking(card: dict, tracking: str, carrier: str):
    entry = card["tracking_entry"]
    combo = card["carrier_combo"]
    if entry.cget("state") == "disabled":
        return
    if tracking:
        entry.delete(0, "end")
        entry.insert(0, tracking)
    if carrier:
        combo.set(carrier)
