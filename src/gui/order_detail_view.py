from __future__ import annotations

import io
import threading
import tkinter as tk
import webbrowser
from tkinter import messagebox

import customtkinter as ctk
import requests
from PIL import Image, ImageTk

from src.ebay_client import EbayClient, EbayOrder
from src.neto_client import NetoClient, NetoOrder


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


class OrderDetailView(ctk.CTkFrame):
    """
    Full-screen order detail view rendered as a plain CTkFrame inside the
    ResultsTab. Replaces the old CTkToplevel modal to avoid canvas rendering bugs.

    Navigation is callback-based:
        on_back()          — return to the results list
        on_fulfilled()     — called after successfully marking order as sent;
                             triggers a list refresh before returning
        on_move_to_unmatched() — optional; moves order out of matched list
    """

    def __init__(
        self,
        master,
        *,
        order_id: str,
        platform: str,
        neto_order: NetoOrder | None,
        ebay_order: EbayOrder | None,
        matched_skus: list[str],
        neto_client: NetoClient | None,
        ebay_client: EbayClient | None,
        dry_run: bool = True,
        on_back,
        on_fulfilled,
        on_move_to_unmatched=None,
        on_book_freight=None,
    ):
        super().__init__(master, fg_color="transparent")
        self._order_id = order_id
        self._platform = platform
        self._neto_order = neto_order
        self._ebay_order = ebay_order
        self._matched_skus = {s.upper().strip() for s in matched_skus}
        self._neto_client = neto_client
        self._ebay_client = ebay_client
        self._dry_run = dry_run
        self._on_back = on_back
        self._on_fulfilled = on_fulfilled
        self._on_move_to_unmatched = on_move_to_unmatched
        self._on_book_freight = on_book_freight
        self._completed = False
        self._image_refs: list = []
        self._full_images: dict[str, bytes] = {}  # url → raw bytes for enlargement
        # Pending image fetches: keyed by api_id (sku for Neto, legacyItemId for eBay)
        self._neto_img_pending: dict[str, ctk.CTkLabel] = {}
        self._ebay_img_pending: dict[str, ctk.CTkLabel] = {}
        self._build_ui()

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── Top navigation bar ────────────────────────────────────────────
        nav = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), corner_radius=0)
        nav.grid(row=0, column=0, sticky="ew")

        self._back_btn = ctk.CTkButton(
            nav, text="← Back to Results", width=150, height=32,
            fg_color="transparent", hover_color=("gray75", "gray30"),
            font=ctk.CTkFont(size=13),
            command=self._on_back,
        )
        self._back_btn.pack(side="left", padx=8, pady=6)

        ctk.CTkLabel(
            nav,
            text=f"Order {self._order_id}  —  {self._platform}",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=8)

        # ── Scrollable body ───────────────────────────────────────────────
        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)

        try:
            self._build_header(body)
            self._build_shipping(body)
            self._build_line_items(body)
            self._build_pricing_summary(body)
            self._build_notes(body)
            self._build_tracking(body)
            self._build_freight_placeholder(body)
            self._build_action_bar(body)
        except Exception as e:
            import traceback
            traceback.print_exc()
            ctk.CTkLabel(
                body, text=f"Error building order detail:\n{e}",
                text_color="red", wraplength=700,
            ).pack(padx=20, pady=20)

    # ── Header ────────────────────────────────────────────────────────────

    def _build_header(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=8, pady=(12, 4))

        ctk.CTkLabel(
            frame,
            text=f"Order {self._order_id}",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left", padx=(0, 6))

        # Copy order number button
        _id_copy_btn = ctk.CTkButton(
            frame, text="Copy", width=45, height=22, font=ctk.CTkFont(size=10),
            fg_color="gray50", hover_color="gray40",
        )
        _id_copy_btn.configure(
            command=lambda b=_id_copy_btn: self._copy_to_clipboard(self._order_id, b)
        )
        _id_copy_btn.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(
            frame,
            text=self._platform,
            font=ctk.CTkFont(size=13),
            fg_color=("dodgerblue3", "dodgerblue4"),
            corner_radius=4,
            text_color="white",
        ).pack(side="left", padx=(0, 12), ipadx=6, ipady=2)

        customer = ""
        date_str = ""
        if self._neto_order:
            customer = self._neto_order.customer_name
            d = self._neto_order.date_paid or self._neto_order.date_placed
            date_str = d.strftime("%d/%m/%Y") if d else ""
        elif self._ebay_order:
            customer = self._ebay_order.buyer_name
            date_str = (
                self._ebay_order.creation_date.strftime("%d/%m/%Y")
                if self._ebay_order.creation_date else ""
            )

        ctk.CTkLabel(frame, text=customer, font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 12))
        ctk.CTkLabel(frame, text=date_str, font=ctk.CTkFont(size=13), text_color="gray50").pack(side="left")

    # ── Shipping Address ──────────────────────────────────────────────────

    def _build_shipping(self, parent):
        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Shipping Details", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        lines = self._get_address_lines()

        for line in lines:
            if not line:
                continue
            row = ctk.CTkFrame(frame, fg_color="transparent")
            row.pack(fill="x", padx=10, pady=1)

            ctk.CTkLabel(row, text=line, font=ctk.CTkFont(size=13), anchor="w").pack(
                side="left", fill="x", expand=True
            )

            btn = ctk.CTkButton(
                row, text="Copy", width=45, height=22, font=ctk.CTkFont(size=10),
                fg_color="gray50", hover_color="gray40",
            )
            btn.configure(command=lambda t=line, b=btn: self._copy_to_clipboard(t, b))
            btn.pack(side="right", padx=4)

        all_text = "\n".join(l for l in lines if l)
        copy_all_btn = ctk.CTkButton(
            frame, text="Copy All", width=70, height=26, font=ctk.CTkFont(size=11),
        )
        copy_all_btn.configure(
            command=lambda b=copy_all_btn: self._copy_to_clipboard(all_text, b, "Copy All")
        )
        copy_all_btn.pack(anchor="e", padx=10, pady=(4, 8))

    def _get_address_lines(self) -> list[str]:
        if self._neto_order:
            o = self._neto_order
            name = f"{o.ship_first_name} {o.ship_last_name}".strip()
            return [
                name,
                o.ship_company,
                o.ship_street1,
                o.ship_street2,
                f"{o.ship_city} {o.ship_state} {o.ship_postcode}".strip(),
                o.ship_country,
                o.ship_phone,
            ]
        elif self._ebay_order:
            o = self._ebay_order
            return [
                o.ship_name,
                o.ship_street1,
                o.ship_street2,
                f"{o.ship_city} {o.ship_state} {o.ship_postcode}".strip(),
                o.ship_country,
                o.ship_phone,
            ]
        return ["(No address available)"]

    # ── Line Items ────────────────────────────────────────────────────────

    def _build_line_items(self, parent):
        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Line Items", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(0, 2))
        for text, w in [("", 54), ("SKU", 130), ("Description", 260), ("Qty", 40), ("Price", 70), ("Arrived", 60)]:
            ctk.CTkLabel(hdr, text=text, width=w, font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(
                side="left", padx=(0, 6)
            )

        items_frame = ctk.CTkFrame(frame, fg_color="transparent")
        items_frame.pack(fill="x", padx=10, pady=(0, 8))

        if self._neto_order:
            self._build_neto_line_items(items_frame)
        elif self._ebay_order:
            self._ebay_note_widgets: list[tuple] = []
            self._build_ebay_line_items(items_frame)

        # Kick off background image fetch now that all labels exist
        if self._neto_img_pending or self._ebay_img_pending:
            threading.Thread(target=self._fetch_product_images, daemon=True).start()

    def _build_neto_line_items(self, items_frame):
        for li in self._neto_order.line_items:
            self._build_line_item_row(items_frame, li.sku, li.product_name,
                                       li.quantity, li.unit_price, li.sku)

    def _build_ebay_line_items(self, items_frame):
        for li in self._ebay_order.line_items:
            self._build_line_item_row(items_frame, li.sku, li.title,
                                       li.quantity, li.unit_price, li.legacy_item_id)
            # Compact inline note editor — single row: label, entry, char count, save btn
            note_row = ctk.CTkFrame(items_frame, fg_color=("gray90", "gray25"), corner_radius=4)
            note_row.pack(fill="x", padx=(54, 0), pady=(0, 6))

            ctk.CTkLabel(
                note_row, text="Notes:", width=50,
                font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
            ).pack(side="left", padx=(6, 4), pady=4)

            entry = ctk.CTkEntry(note_row, font=ctk.CTkFont(size=11), height=26, text_color="#f5c518")
            entry.pack(side="left", fill="x", expand=True, pady=4)
            if li.notes:
                entry.insert(0, li.notes)
            self._bind_context_menu(entry._entry)

            char_label = ctk.CTkLabel(
                note_row, text=f"{len(li.notes)}/255", width=50,
                font=ctk.CTkFont(size=10), text_color="gray50",
            )
            char_label.pack(side="left", padx=(4, 0), pady=4)
            entry.bind("<KeyRelease>", lambda e, en=entry, cl=char_label: self._update_char_count_entry(en, cl))

            btn = ctk.CTkButton(
                note_row, text="Save", width=55, height=24,
                font=ctk.CTkFont(size=11),
                command=lambda l=li, en=entry, b=None: self._save_ebay_note_entry(l, en, b),
            )
            btn.pack(side="left", padx=(4, 6), pady=4)
            btn.configure(command=lambda l=li, en=entry, b=btn: self._save_ebay_note_entry(l, en, b))

            self._ebay_note_widgets.append((li, entry, btn))

    def _build_line_item_row(self, items_frame, sku, desc, qty, price, api_id):
        row = ctk.CTkFrame(items_frame, fg_color="transparent")
        row.pack(fill="x", pady=1)

        img_label = ctk.CTkLabel(row, text="·", width=50, height=50,
                                 text_color="gray50", font=ctk.CTkFont(size=20))
        img_label.pack(side="left", padx=(0, 4))

        # Register for background image fetch
        if api_id:
            if self._neto_order:
                self._neto_img_pending[api_id] = img_label
            elif self._ebay_order:
                self._ebay_img_pending[api_id] = img_label

        arrived = "✓" if sku.upper().strip() in self._matched_skus else ""
        price_str = f"${price:.2f}" if price else ""
        ctk.CTkLabel(row, text=sku, width=130, anchor="w", wraplength=130).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(row, text=desc, width=260, anchor="w", wraplength=260).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(row, text=str(qty), width=40, anchor="w").pack(side="left", padx=(0, 6))
        ctk.CTkLabel(row, text=price_str, width=70, anchor="w").pack(side="left", padx=(0, 6))
        ctk.CTkLabel(
            row, text=arrived, width=60, anchor="w",
            text_color="green" if arrived else "gray50",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")

    # ── Pricing Summary ──────────────────────────────────────────────────

    def _build_pricing_summary(self, parent):
        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Pricing & Shipping", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        details = ctk.CTkFrame(frame, fg_color="transparent")
        details.pack(fill="x", padx=10, pady=(0, 8))

        if self._neto_order:
            o = self._neto_order
            shipping_total = o.shipping_total
            grand_total = o.grand_total
            shipping_method = o.shipping_method
            shipping_type = o.shipping_type
        elif self._ebay_order:
            o = self._ebay_order
            shipping_total = o.shipping_cost
            grand_total = o.order_total
            shipping_method = o.shipping_method
            shipping_type = o.shipping_type
        else:
            return

        rows = []
        if shipping_method:
            rows.append(("Shipping Method:", shipping_method))
        if shipping_type:
            rows.append(("Shipping Type:", shipping_type))
        rows.append(("Shipping Cost:", f"${shipping_total:.2f}"))
        rows.append(("Order Total:", f"${grand_total:.2f}"))

        for label_text, value_text in rows:
            row = ctk.CTkFrame(details, fg_color="transparent")
            row.pack(fill="x", pady=1)
            ctk.CTkLabel(row, text=label_text, width=130, anchor="w",
                         font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")

            # Highlight shipping type with a colored tag
            if label_text == "Shipping Type:":
                tag_colors = {
                    "Express": ("#e74c3c", "white"),
                    "Regular": ("#3498db", "white"),
                    "Local Pickup": ("#2ecc71", "white"),
                }
                fg, tc = tag_colors.get(value_text, ("gray50", "white"))
                ctk.CTkLabel(
                    row, text=value_text, font=ctk.CTkFont(size=12, weight="bold"),
                    fg_color=fg, text_color=tc, corner_radius=4,
                ).pack(side="left", ipadx=6, ipady=1)
            else:
                ctk.CTkLabel(row, text=value_text, anchor="w",
                             font=ctk.CTkFont(size=12)).pack(side="left")

    def _fetch_product_images(self):
        """Background thread: resolve image URLs via API, then download and display."""
        if self._neto_img_pending and self._neto_client:
            try:
                url_map = self._neto_client.get_product_images(list(self._neto_img_pending))
                for sku, url in url_map.items():
                    label = self._neto_img_pending.get(sku)
                    if url and label:
                        self._download_and_show_image(url, label)
            except Exception as e:
                print(f"[IMAGE] Neto image fetch failed: {e}")

        if self._ebay_img_pending and self._ebay_client:
            try:
                url_map = self._ebay_client.get_item_images(list(self._ebay_img_pending))
                for item_id, url in url_map.items():
                    label = self._ebay_img_pending.get(item_id)
                    if url and label:
                        self._download_and_show_image(url, label)
            except Exception as e:
                print(f"[IMAGE] eBay image fetch failed: {e}")

    def _download_and_show_image(self, url: str, label: ctk.CTkLabel):
        """
        Download an image synchronously (call from a background thread) and
        schedule a UI update on the main thread. Also enables click-to-enlarge.
        """
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
            print(f"[IMAGE] Failed to download {url!r}: {exc}")

    def _load_image_async(self, url: str, label: ctk.CTkLabel):
        """Spawn a thread to download and display an image (for use from main thread)."""
        threading.Thread(
            target=self._download_and_show_image, args=(url, label), daemon=True
        ).start()

    def _open_image_large(self, url: str):
        """Open a plain tk.Toplevel showing a larger version of the item image."""
        data = self._full_images.get(url)
        if not data:
            return  # Still loading or failed — no-op

        try:
            img = Image.open(io.BytesIO(data))
            img.thumbnail((600, 600), Image.Resampling.LANCZOS)

            # Use plain tk.Toplevel (not CTkToplevel) to avoid canvas rendering bugs
            top = tk.Toplevel(self.winfo_toplevel())
            top.title("Image Preview")
            top.resizable(False, False)

            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(top, image=photo, cursor="hand2")
            lbl.image = photo  # Prevent garbage collection
            lbl.pack()

            hint = tk.Label(top, text="Click image or press Escape to close", fg="gray60")
            hint.pack(pady=(0, 4))

            lbl.bind("<Button-1>", lambda e: top.destroy())
            top.bind("<Escape>", lambda e: top.destroy())

            # Center on screen
            top.update_idletasks()
            w, h = top.winfo_width(), top.winfo_height()
            sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
            top.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

        except Exception:
            pass

    # ── Notes ─────────────────────────────────────────────────────────────

    def _build_notes(self, parent):
        if self._ebay_order:
            # Bottom notes section disabled — per-item notes are shown inline in
            # _build_ebay_line_items, making this compiled summary redundant.
            # Uncomment to restore:
            # notes = self._ebay_order.buyer_notes
            # if not notes:
            #     return
            # frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
            # frame.pack(fill="x", padx=8, pady=6)
            # ctk.CTkLabel(frame, text="Notes", font=ctk.CTkFont(size=13, weight="bold")).pack(
            #     anchor="w", padx=10, pady=(8, 4)
            # )
            # ctk.CTkLabel(
            #     frame, text=f"Buyer notes: {notes}",
            #     font=ctk.CTkFont(size=12), anchor="w", wraplength=700,
            #     text_color="#f5c518",
            # ).pack(fill="x", padx=10, pady=(0, 8))
            return

        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Notes", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        # Keep a reference to the outer frame so _rebuild_notes_content can recreate
        # the inner content in-place without the outer frame losing its pack position.
        self._notes_content_parent = frame
        self._notes_content_frame = ctk.CTkFrame(frame, fg_color="transparent")
        self._notes_content_frame.pack(fill="x")
        self._build_neto_notes(self._notes_content_frame)

    def _build_neto_notes(self, parent):
        o = self._neto_order

        if o.delivery_instruction:
            ctk.CTkLabel(
                parent, text=f"Delivery Instructions: {o.delivery_instruction}",
                font=ctk.CTkFont(size=12), anchor="w", wraplength=700,
                text_color="#f5c518",
            ).pack(fill="x", padx=10, pady=2)

        if o.sticky_notes:
            ctk.CTkLabel(parent, text="Existing Sticky Notes:", font=ctk.CTkFont(size=12, weight="bold")).pack(
                anchor="w", padx=10, pady=(6, 2)
            )
            # Sort newest first by StickyNoteID (higher ID = more recently created)
            def _note_sort_key(n):
                try:
                    return int(n.get("StickyNoteID") or 0)
                except (ValueError, TypeError):
                    return 0
            for note in sorted(o.sticky_notes, key=_note_sort_key, reverse=True):
                title = note.get("Title", "")
                desc = note.get("Description", "")
                # DateAdded is the field Neto uses; fall back to other common names
                date_str = (
                    note.get("DateAdded") or note.get("DateCreated") or
                    note.get("CreatedDate") or ""
                )
                header = f"[{date_str}] " if date_str else ""
                text = f"{header}{title}: {desc}" if title else f"{header}{desc}"
                tb = ctk.CTkTextbox(
                    parent, height=50, font=ctk.CTkFont(size=12),
                    text_color="#f5c518", fg_color=("gray90", "gray25"),
                    border_width=0,
                )
                tb.insert("1.0", text)
                tb.configure(state="disabled")
                tb.pack(fill="x", padx=10, pady=1)

        if o.internal_notes:
            ctk.CTkLabel(
                parent, text=f"Internal Notes: {o.internal_notes}",
                font=ctk.CTkFont(size=12), anchor="w", wraplength=700,
                text_color="#f5c518",
            ).pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(parent, text="Add Sticky Note:", font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=10, pady=(8, 2)
        )
        self._note_textbox = ctk.CTkTextbox(parent, height=60, font=ctk.CTkFont(size=12))
        self._note_textbox.pack(fill="x", padx=10, pady=(0, 4))
        self._bind_context_menu(self._note_textbox._textbox)

        self._add_note_btn = ctk.CTkButton(
            parent, text="Add Note", width=90, height=28,
            command=self._add_neto_note,
        )
        self._add_note_btn.pack(anchor="e", padx=10, pady=(0, 8))

    def _update_char_count(self, textbox, label):
        text = textbox.get("1.0", "end").strip()
        count = len(text)
        color = "red" if count > 255 else "gray50"
        label.configure(text=f"{count}/255", text_color=color)

    def _update_char_count_entry(self, entry, label):
        text = entry.get().strip()
        count = len(text)
        color = "red" if count > 255 else "gray50"
        label.configure(text=f"{count}/255", text_color=color)

    def _save_ebay_note_entry(self, line_item, entry, btn):
        text = entry.get().strip()
        if not text:
            return
        if len(text) > 255:
            messagebox.showwarning(
                "Too Long",
                f"eBay PrivateNotes are limited to 255 characters.\nCurrent length: {len(text)}",
                parent=self.winfo_toplevel(),
            )
            return
        if not line_item.legacy_item_id or not line_item.legacy_transaction_id:
            messagebox.showerror(
                "Error",
                "Cannot save note: missing eBay item/transaction IDs for this line item.",
                parent=self.winfo_toplevel(),
            )
            return
        try:
            self._ebay_client.set_private_notes(
                item_id=line_item.legacy_item_id,
                transaction_id=line_item.legacy_transaction_id,
                note_text=text,
                dry_run=self._dry_run,
            )
            line_item.notes = text
            parent = self.winfo_toplevel()
            if self._dry_run:
                if btn:
                    btn.configure(state="disabled", text="Saved")
                    self.after(2000, lambda b=btn: b.configure(state="normal", text="Save"))
                messagebox.showinfo("Dry Run", f"[DRY RUN] PrivateNotes would be set:\n{text}", parent=parent)
            else:
                if btn:
                    btn.configure(state="disabled", text="Refreshing…")

                def _fetch():
                    try:
                        fresh = self._ebay_client.get_orders_by_ids([self._order_id])
                        self.after(0, lambda: self._on_ebay_note_refreshed(fresh, btn))
                    except Exception as exc:
                        self.after(0, lambda err=str(exc), b=btn: self._on_ebay_note_refresh_failed(err, b))

                threading.Thread(target=_fetch, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save note:\n{e}", parent=self.winfo_toplevel())

    def _add_neto_note(self):
        from datetime import date
        text = self._note_textbox.get("1.0", "end").strip()
        if not text:
            return
        dated_text = f"[{date.today().strftime('%d/%m/%Y')}] {text}"
        try:
            self._neto_client.add_sticky_note(
                self._order_id,
                title="Packing Note",
                description=dated_text,
                dry_run=self._dry_run,
            )
            parent = self.winfo_toplevel()
            if self._dry_run:
                self._add_note_btn.configure(state="disabled", text="Note Added")
                messagebox.showinfo("Dry Run", f"[DRY RUN] Sticky note would be added:\n{dated_text}", parent=parent)
            else:
                # Re-fetch this order so the sticky notes list refreshes in-place.
                self._add_note_btn.configure(state="disabled", text="Refreshing…")

                def _fetch():
                    try:
                        fresh = self._neto_client.get_orders_by_ids([self._order_id])
                        self.after(0, lambda: self._on_note_refreshed(fresh))
                    except Exception as exc:
                        self.after(0, lambda err=str(exc): self._on_note_refresh_failed(err))

                threading.Thread(target=_fetch, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add note: {e}", parent=self.winfo_toplevel())

    def _on_note_refreshed(self, fresh_orders: list):
        """Called on the main thread after the post-save re-fetch completes."""
        if fresh_orders:
            self._neto_order = fresh_orders[0]
        self._rebuild_notes_content()

    def _on_note_refresh_failed(self, error_msg: str):
        """Called when the post-save re-fetch fails; note was saved but display is stale."""
        self._add_note_btn.configure(state="disabled", text="Note Added")
        self._status_label.configure(
            text="Note saved, but couldn't refresh display.", text_color="orange"
        )

    def _on_ebay_note_refreshed(self, fresh_orders: list, btn):
        """Called on the main thread after eBay re-fetch completes post note-save."""
        if fresh_orders:
            self._ebay_order = fresh_orders[0]
            fresh_items = {
                (fli.legacy_item_id, fli.legacy_transaction_id): fli
                for fli in self._ebay_order.line_items
            }
            for li, entry_widget, _btn in self._ebay_note_widgets:
                fli = fresh_items.get((li.legacy_item_id, li.legacy_transaction_id))
                if fli:
                    li.notes = fli.notes
                    li.legacy_transaction_id = fli.legacy_transaction_id
                    entry_widget.delete(0, "end")
                    entry_widget.insert(0, fli.notes or "")
        if btn:
            btn.configure(state="normal", text="Save")

    def _on_ebay_note_refresh_failed(self, error_msg: str, btn):
        """Called when eBay re-fetch fails; note was saved but entries may be stale."""
        if btn:
            btn.configure(state="disabled", text="Saved")
            self.after(2000, lambda b=btn: b.configure(state="normal", text="Save"))

    def _rebuild_notes_content(self):
        """Destroy and recreate the notes content area inside the outer notes frame.

        The outer frame keeps its pack position so the section doesn't jump to
        the bottom of the scrollable view.
        """
        self._notes_content_frame.destroy()
        self._notes_content_frame = ctk.CTkFrame(self._notes_content_parent, fg_color="transparent")
        self._notes_content_frame.pack(fill="x")
        self._build_neto_notes(self._notes_content_frame)

    # ── Tracking ──────────────────────────────────────────────────────────

    def _build_tracking(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Tracking", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(0, 4)
        )

        row = ctk.CTkFrame(frame, fg_color="transparent")
        row.pack(fill="x", padx=10)

        ctk.CTkLabel(row, text="Tracking Number:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 6))
        self._tracking_entry = ctk.CTkEntry(row, width=220, font=ctk.CTkFont(size=12))
        self._tracking_entry.pack(side="left", padx=(0, 20))
        self._bind_context_menu(self._tracking_entry._entry)

        ctk.CTkLabel(row, text="Shipping Method:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 6))
        self._carrier_combo = ctk.CTkComboBox(
            row, values=_SHIPPING_METHODS, width=180, font=ctk.CTkFont(size=12),
        )
        self._carrier_combo.set("")  # Blank default
        self._carrier_combo.pack(side="left")

    # ── Freight Placeholder ───────────────────────────────────────────────

    def _build_freight_placeholder(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=8, pady=6)

        if self._on_book_freight:
            ctk.CTkButton(
                frame, text="Book Freight", width=120, height=30,
                fg_color=("dodgerblue", "dodgerblue3"),
                hover_color=("dodgerblue3", "dodgerblue"),
                command=lambda: self._on_book_freight(self._order_id, self._platform),
            ).pack(side="left", padx=10)
        else:
            ctk.CTkButton(
                frame, text="Book Freight", width=120, height=30, state="disabled",
                fg_color="gray50",
            ).pack(side="left", padx=10)
            ctk.CTkLabel(frame, text="Shipping not configured", font=ctk.CTkFont(size=11),
                         text_color="gray50").pack(side="left")

    # ── Action Bar ────────────────────────────────────────────────────────

    def _build_action_bar(self, parent):
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.pack(fill="x", padx=8, pady=(8, 16))

        self._send_btn = ctk.CTkButton(
            bar, text="Mark as Sent", width=140, height=36,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
            command=self._mark_as_sent,
        )
        self._send_btn.pack(side="left", padx=(0, 12))

        if self._on_move_to_unmatched:
            ctk.CTkButton(
                bar, text="Move to Unmatched", width=150, height=36,
                fg_color="gray50", hover_color="gray40",
                command=self._do_move_to_unmatched,
            ).pack(side="left", padx=(0, 12))

        ctk.CTkButton(
            bar, text="← Back", width=90, height=36,
            fg_color="gray50", hover_color="gray40",
            command=self._on_back,
        ).pack(side="left")

        self._status_label = ctk.CTkLabel(bar, text="", font=ctk.CTkFont(size=13))
        self._status_label.pack(side="left", padx=12)

    # ── Actions ───────────────────────────────────────────────────────────

    def _mark_as_sent(self):
        tracking = self._tracking_entry.get().strip()
        carrier = self._carrier_combo.get().strip()
        parent = self.winfo_toplevel()

        if not tracking:
            if not messagebox.askyesno("No Tracking", "Send without a tracking number?", parent=parent):
                return

        try:
            if self._neto_order and self._neto_client:
                line_item_skus = [li.sku for li in self._neto_order.line_items]
                self._neto_client.update_order_status(
                    self._order_id,
                    new_status="Dispatched",
                    tracking_number=tracking,
                    shipping_method=carrier,
                    line_item_skus=line_item_skus,
                    dry_run=self._dry_run,
                )
            elif self._ebay_order and self._ebay_client:
                self._ebay_client.create_shipping_fulfillment(
                    self._order_id,
                    line_items=self._ebay_order.line_items,
                    tracking_number=tracking,
                    carrier=carrier,
                    dry_run=self._dry_run,
                )

            self._completed = True
            self._send_btn.configure(state="disabled", text="SENT", fg_color="gray50")
            self._tracking_entry.configure(state="disabled")
            self._carrier_combo.configure(state="disabled")

            if self._dry_run:
                self._status_label.configure(text="[DRY RUN] Order marked as sent", text_color="orange")
            else:
                self._status_label.configure(text="Order marked as sent!", text_color="green")
                # Kogan orders are tracked via Neto but must be manually dispatched
                # on the Kogan portal — open it automatically after marking as sent
                if (
                    self._neto_order
                    and self._neto_order.sales_channel.lower() == "kogan"
                ):
                    webbrowser.open("https://dispatch.aws.kgn.io/Manage")
                self._on_fulfilled()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to mark order as sent:\n{e}", parent=parent)

    def _do_move_to_unmatched(self):
        if self._on_move_to_unmatched:
            self._on_move_to_unmatched()
        self._on_back()

    def set_tracking(self, tracking: str = "", carrier: str = ""):
        """Programmatically fill the tracking entry and carrier combobox."""
        if tracking:
            self._tracking_entry.delete(0, "end")
            self._tracking_entry.insert(0, tracking)
        if carrier:
            self._carrier_combo.set(carrier)

    def show_completed_warning(self):
        """Call from ResultsTab after status check confirms order is already done."""
        self._status_label.configure(
            text="⚠ This order has already been completed", text_color="orange"
        )
        self._send_btn.configure(state="disabled")

    # ── Helpers ───────────────────────────────────────────────────────────

    def _bind_context_menu(self, widget):
        """Bind a right-click Cut/Copy/Paste context menu to a tk widget."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Cut",   command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Copy",  command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
        widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

    def _copy_to_clipboard(self, text: str, btn=None, original_text: str = "Copy"):
        """Copy text to clipboard; briefly change btn label to 'Copied!' if provided."""
        self.clipboard_clear()
        self.clipboard_append(text)
        if btn is not None:
            btn.configure(text="Copied!")
            self.after(2000, lambda: btn.configure(text=original_text))
