from __future__ import annotations

import io
import threading
from tkinter import messagebox

import customtkinter as ctk
import requests
from PIL import Image

from src.ebay_client import EbayClient, EbayLineItem, EbayOrder
from src.neto_client import NetoClient, NetoOrder


_PLACEHOLDER_SIZE = (50, 50)


class OrderDetailModal(ctk.CTkToplevel):
    """Modal window showing full order details with fulfillment actions."""

    def __init__(
        self,
        master,
        order_id: str,
        platform: str,
        neto_order: NetoOrder | None,
        ebay_order: EbayOrder | None,
        matched_skus: list[str],
        neto_client: NetoClient | None,
        ebay_client: EbayClient | None,
        dry_run: bool = True,
        on_close_callback=None,
    ):
        super().__init__(master)

        # Workaround for customtkinter bug on Windows: CTkToplevel schedules
        # after(200, iconbitmap(...)) which fails with TclError on some systems.
        _orig_iconbitmap = self.iconbitmap
        def _safe_iconbitmap(*args, **kwargs):
            try:
                _orig_iconbitmap(*args, **kwargs)
            except Exception:
                pass
        self.iconbitmap = _safe_iconbitmap

        self._order_id = order_id
        self._platform = platform
        self._neto_order = neto_order
        self._ebay_order = ebay_order
        self._matched_skus = set(s.upper().strip() for s in matched_skus)
        self._neto_client = neto_client
        self._ebay_client = ebay_client
        self._dry_run = dry_run
        self._on_close = on_close_callback
        self._completed = False
        self._image_refs: list = []  # prevent GC of CTkImage objects

        self.title(f"Order {order_id} — {platform}")
        self.geometry("750x780")
        self.minsize(600, 500)

        # Use the top-level window as transient parent (not the tab frame)
        toplevel = master.winfo_toplevel()
        self.transient(toplevel)

        self.protocol("WM_DELETE_WINDOW", self._close)

        # Build UI first, then grab focus after a short delay to avoid Windows freeze
        self._build_ui()
        self.after(150, self._activate)

    def _build_ui(self):
        try:
            container = ctk.CTkScrollableFrame(self, fg_color="transparent")
            container.pack(fill="both", expand=True, padx=8, pady=8)

            self._build_header(container)
            self._build_shipping(container)
            self._build_line_items(container)
            self._build_notes(container)
            self._build_tracking(container)
            self._build_freight_placeholder(container)
            self._build_action_bar()
        except Exception as e:
            import traceback
            traceback.print_exc()
            ctk.CTkLabel(
                self, text=f"Error building order detail:\n{e}",
                text_color="red", wraplength=600,
            ).pack(padx=20, pady=20)

    # ── Header ────────────────────────────────────────────────────────────

    def _build_header(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=8, pady=(8, 4))

        ctk.CTkLabel(
            frame,
            text=f"Order {self._order_id}",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left", padx=(0, 12))

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
            date_str = d.strftime("%Y-%m-%d") if d else ""
        elif self._ebay_order:
            customer = self._ebay_order.buyer_name
            date_str = self._ebay_order.creation_date.strftime("%Y-%m-%d") if self._ebay_order.creation_date else ""

        ctk.CTkLabel(frame, text=customer, font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 12))
        ctk.CTkLabel(frame, text=date_str, font=ctk.CTkFont(size=13), text_color="gray50").pack(side="left")

    # ── Shipping Address ──────────────────────────────────────────────────

    def _build_shipping(self, parent):
        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Shipping Address", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        lines = self._get_address_lines()
        self._address_lines = lines

        for line in lines:
            if not line:
                continue
            row = ctk.CTkFrame(frame, fg_color="transparent")
            row.pack(fill="x", padx=10, pady=1)
            ctk.CTkLabel(row, text=line, font=ctk.CTkFont(size=13), anchor="w").pack(side="left", fill="x", expand=True)
            ctk.CTkButton(
                row, text="Copy", width=45, height=22, font=ctk.CTkFont(size=10),
                fg_color="gray50", hover_color="gray40",
                command=lambda t=line: self._copy_to_clipboard(t),
            ).pack(side="right", padx=4)

        ctk.CTkButton(
            frame, text="Copy All", width=70, height=26, font=ctk.CTkFont(size=11),
            command=lambda: self._copy_to_clipboard("\n".join(l for l in lines if l)),
        ).pack(anchor="e", padx=10, pady=(4, 8))

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

        # Header row
        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(0, 2))
        for text, w in [("", 54), ("SKU", 130), ("Description", 250), ("Qty", 40), ("Arrived", 50)]:
            ctk.CTkLabel(hdr, text=text, width=w, font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(
                side="left", padx=(0, 6)
            )

        items_frame = ctk.CTkFrame(frame, fg_color="transparent")
        items_frame.pack(fill="x", padx=10, pady=(0, 8))

        line_items = []
        if self._neto_order:
            for li in self._neto_order.line_items:
                line_items.append((li.sku, li.product_name, li.quantity, li.image_url))
        elif self._ebay_order:
            for li in self._ebay_order.line_items:
                line_items.append((li.sku, li.title, li.quantity, li.image_url))

        for sku, desc, qty, image_url in line_items:
            row = ctk.CTkFrame(items_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)

            # Image placeholder
            img_label = ctk.CTkLabel(row, text="", width=50, height=50)
            img_label.pack(side="left", padx=(0, 4))

            if image_url:
                self._load_image_async(image_url, img_label)

            arrived = "*" if sku.upper().strip() in self._matched_skus else ""

            ctk.CTkLabel(row, text=sku, width=130, anchor="w", wraplength=130).pack(side="left", padx=(0, 6))
            ctk.CTkLabel(row, text=desc, width=250, anchor="w", wraplength=250).pack(side="left", padx=(0, 6))
            ctk.CTkLabel(row, text=str(qty), width=40, anchor="w").pack(side="left", padx=(0, 6))
            ctk.CTkLabel(
                row, text=arrived, width=50, anchor="w",
                text_color="green" if arrived else "gray50",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(side="left")

    def _load_image_async(self, url: str, label: ctk.CTkLabel):
        """Load a product image in a background thread."""
        def _fetch():
            try:
                resp = requests.get(url, timeout=10)
                resp.raise_for_status()
                img = Image.open(io.BytesIO(resp.content))
                img.thumbnail(_PLACEHOLDER_SIZE, Image.Resampling.LANCZOS)
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=_PLACEHOLDER_SIZE)
                self._image_refs.append(ctk_img)
                label.after(0, lambda: label.configure(image=ctk_img, text=""))
            except Exception:
                pass  # leave placeholder

        threading.Thread(target=_fetch, daemon=True).start()

    # ── Notes ─────────────────────────────────────────────────────────────

    def _build_notes(self, parent):
        frame = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(frame, text="Notes", font=ctk.CTkFont(size=13, weight="bold")).pack(
            anchor="w", padx=10, pady=(8, 4)
        )

        if self._neto_order:
            self._build_neto_notes(frame)
        elif self._ebay_order:
            self._build_ebay_notes(frame)

    def _build_neto_notes(self, parent):
        o = self._neto_order

        # Delivery instructions (read-only)
        if o.delivery_instruction:
            ctk.CTkLabel(
                parent, text=f"Delivery Instructions: {o.delivery_instruction}",
                font=ctk.CTkFont(size=12), anchor="w", wraplength=650,
            ).pack(fill="x", padx=10, pady=2)

        # Existing sticky notes (read-only)
        if o.sticky_notes:
            ctk.CTkLabel(parent, text="Existing Sticky Notes:", font=ctk.CTkFont(size=12, weight="bold")).pack(
                anchor="w", padx=10, pady=(6, 2)
            )
            for note in o.sticky_notes:
                title = note.get("Title", "")
                desc = note.get("Description", "")
                text = f"{title}: {desc}" if title else desc
                ctk.CTkLabel(
                    parent, text=text, font=ctk.CTkFont(size=12),
                    anchor="w", wraplength=650, fg_color=("gray90", "gray25"),
                    corner_radius=4,
                ).pack(fill="x", padx=10, pady=1, ipadx=4, ipady=2)

        # Internal notes (read-only)
        if o.internal_notes:
            ctk.CTkLabel(
                parent, text=f"Internal Notes: {o.internal_notes}",
                font=ctk.CTkFont(size=12), anchor="w", wraplength=650,
            ).pack(fill="x", padx=10, pady=2)

        # New sticky note
        ctk.CTkLabel(parent, text="Add Sticky Note:", font=ctk.CTkFont(size=12)).pack(
            anchor="w", padx=10, pady=(8, 2)
        )
        self._note_textbox = ctk.CTkTextbox(parent, height=60, font=ctk.CTkFont(size=12))
        self._note_textbox.pack(fill="x", padx=10, pady=(0, 4))

        self._add_note_btn = ctk.CTkButton(
            parent, text="Add Note", width=90, height=28,
            command=self._add_neto_note,
        )
        self._add_note_btn.pack(anchor="e", padx=10, pady=(0, 8))

    def _build_ebay_notes(self, parent):
        o = self._ebay_order

        # Buyer checkout notes (read-only)
        if o.buyer_notes:
            ctk.CTkLabel(
                parent, text=f"Buyer Notes: {o.buyer_notes}",
                font=ctk.CTkFont(size=12), anchor="w", wraplength=650,
            ).pack(fill="x", padx=10, pady=2)

        # Per-item PrivateNotes
        for li in o.line_items:
            if li.notes:
                ctk.CTkLabel(
                    parent, text=f"[{li.sku}] PrivateNotes: {li.notes}",
                    font=ctk.CTkFont(size=12), anchor="w", wraplength=650,
                    fg_color=("gray90", "gray25"), corner_radius=4,
                ).pack(fill="x", padx=10, pady=1, ipadx=4, ipady=2)

        # Placeholder for note editing (eBay PrivateNotes editing is complex)
        ctk.CTkLabel(
            parent,
            text="(eBay note editing coming in a future update)",
            font=ctk.CTkFont(size=11), text_color="gray50",
        ).pack(anchor="w", padx=10, pady=(4, 8))

    def _add_neto_note(self):
        text = self._note_textbox.get("1.0", "end").strip()
        if not text:
            return
        try:
            self._neto_client.add_sticky_note(
                self._order_id,
                title="Packing Note",
                description=text,
                dry_run=self._dry_run,
            )
            self._add_note_btn.configure(state="disabled", text="Note Added")
            if self._dry_run:
                messagebox.showinfo("Dry Run", f"[DRY RUN] Sticky note would be added:\n{text}", parent=self)
            else:
                messagebox.showinfo("Success", "Sticky note added.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add note: {e}", parent=self)

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

        ctk.CTkLabel(row, text="Carrier:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 6))
        self._carrier_entry = ctk.CTkEntry(row, width=160, font=ctk.CTkFont(size=12))
        self._carrier_entry.pack(side="left")

    # ── Freight Placeholder ───────────────────────────────────────────────

    def _build_freight_placeholder(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=8, pady=6)

        ctk.CTkButton(
            frame, text="Book Freight", width=120, height=30, state="disabled",
            fg_color="gray50",
        ).pack(side="left", padx=10)
        ctk.CTkLabel(frame, text="Coming soon", font=ctk.CTkFont(size=11), text_color="gray50").pack(side="left")

    # ── Action Bar ────────────────────────────────────────────────────────

    def _build_action_bar(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=12, pady=(4, 12))

        self._send_btn = ctk.CTkButton(
            bar, text="Mark as Sent", width=140, height=36,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
            command=self._mark_as_sent,
        )
        self._send_btn.pack(side="left", padx=(0, 12))

        self._close_btn = ctk.CTkButton(
            bar, text="Close", width=90, height=36,
            fg_color="gray50", hover_color="gray40",
            command=self._close,
        )
        self._close_btn.pack(side="left")

        self._status_label = ctk.CTkLabel(bar, text="", font=ctk.CTkFont(size=13))
        self._status_label.pack(side="left", padx=12)

    def _mark_as_sent(self):
        tracking = self._tracking_entry.get().strip()
        carrier = self._carrier_entry.get().strip()

        if not tracking:
            if not messagebox.askyesno(
                "No Tracking",
                "Send without a tracking number?",
                parent=self,
            ):
                return

        try:
            if self._neto_order and self._neto_client:
                self._neto_client.update_order_status(
                    self._order_id,
                    new_status="Dispatched",
                    tracking_number=tracking,
                    carrier=carrier,
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
            self._send_btn.configure(state="disabled", text="COMPLETED", fg_color="gray50")
            self._tracking_entry.configure(state="disabled")
            self._carrier_entry.configure(state="disabled")
            self._close_btn.configure(text="Back to Orders")

            if self._dry_run:
                self._status_label.configure(text="[DRY RUN] Order marked as sent", text_color="orange")
            else:
                self._status_label.configure(text="Order marked as sent!", text_color="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to mark order as sent:\n{e}", parent=self)

    def _activate(self):
        """Bring window to front and grab focus after UI is fully rendered."""
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except Exception:
            pass  # grab can fail if window was closed quickly

    def _close(self):
        try:
            self.grab_release()
        except Exception:
            pass
        if self._on_close:
            self._on_close(self._completed)
        self.destroy()

    def _copy_to_clipboard(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
