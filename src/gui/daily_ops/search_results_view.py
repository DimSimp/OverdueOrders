from __future__ import annotations

import logging
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

import customtkinter as ctk

from src.gui.results_tab import OrderTreeview
from src.ebay_client import EbayAuthError, EbayAPIError
from src.neto_client import NetoAPIError
from src.gui.daily_ops.search_options_view import (
    _filter_neto_by_channel,
    _apply_text_filters,
    _NETO_ORDER_RE,
    _EBAY_ORDER_RE,
)

log = logging.getLogger(__name__)

# ── Column spec (same as AllOrdersResultsView) ─────────────────────────────────

_SEARCH_COL_SPEC = {
    "#0":          ("Order No.",   120),
    "date":        ("Date",         70),
    "platform":    ("Platform",     80),
    "customer":    ("Customer",    170),
    "shipping":    ("Shipping",     90),
    "assigned":    ("Assigned",     80),
    "sku":         ("SKU",         130),
    "description": ("Description", 180),
    "qty":         ("Qty",          40),
    "notes":       ("Value",       100),
    "status":      ("Status",       90),
    "order_notes": ("Notes",       150),
}

# Maps eBay API fulfillment status → human-readable display text
_EBAY_STATUS_DISPLAY = {
    "NOT_STARTED": "Awaiting Dispatch",
    "IN_PROGRESS":  "In Progress",
    "FULFILLED":    "Dispatched",
}


class SearchResultsView(ctk.CTkFrame):
    """
    Results & Dispatch view for Search for Order.

    Shows search results in a single flat list. Refresh re-runs the original
    search query. No session persistence, no pick list export.
    """

    def __init__(self, master, window, search_options: dict, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._search_options = search_options
        self._on_back = on_back
        self._assign_filter = "All"

        self._neto_orders: list = []
        self._ebay_orders: list = []

        self._hidden_order_ids: set[tuple] = set()

        self._detail_frame = None
        self._freight_frame = None
        self._last_clicked_order_id: str | None = None
        self._last_clicked_platform: str | None = None

        self._collated_groups: dict = {}
        self._collated_frame = None
        self._ungrouped_order_ids: set[str] = set()
        self._book_freight_for_all: bool = False

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_ui()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        self._list_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._list_frame.grid(row=0, column=0, sticky="nsew")
        self._list_frame.grid_rowconfigure(1, weight=1)
        self._list_frame.grid_rowconfigure(2, weight=0)
        self._list_frame.grid_columnconfigure(0, weight=1)

        self._build_list_page(self._list_frame)

    def _build_list_page(self, parent):
        # ── Toolbar ─────────────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(parent, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))

        self._back_btn = ctk.CTkButton(
            toolbar, text="← Back to Search", width=140,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=self._on_back,
        )
        self._back_btn.pack(side="left", padx=(0, 6))

        self._refresh_btn = ctk.CTkButton(
            toolbar, text="Refresh", width=90,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._refresh_all_orders,
        )
        self._refresh_btn.pack(side="left", padx=(0, 6))

        self._cancel_btn = ctk.CTkButton(
            toolbar, text="Cancel / Reprint", width=140,
            fg_color=("firebrick3", "firebrick4"),
            hover_color=("firebrick4", "firebrick"),
            command=self._open_cancel_dialog,
        )
        self._cancel_btn.pack(side="left", padx=(0, 10))

        self._status_label = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        )
        self._status_label.pack(side="right")

        self._error_label = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(size=12), text_color="red",
        )
        self._error_label.pack(side="right", padx=(0, 8))

        self._assign_seg = ctk.CTkSegmentedButton(
            toolbar, values=["All", "Mine", "Unassigned"],
            command=self._on_assign_filter_change,
            height=28, font=ctk.CTkFont(size=12),
        )
        self._assign_seg.set("All")
        self._assign_seg.pack(side="right", padx=(0, 12))

        ctk.CTkLabel(
            toolbar, text="Assign:", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=(0, 4))

        # ── Order list ──────────────────────────────────────────────────────
        list_container = ctk.CTkFrame(parent, fg_color="transparent")
        list_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 4))
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)

        self._tree = OrderTreeview(
            list_container,
            col_spec=_SEARCH_COL_SPEC,
            table_id="daily_search_results",
            user_manager=getattr(self._window, "user_manager", None),
            current_user=getattr(self._window, "current_user", None),
            on_row_click=self._open_detail_view,
            on_context_action=self._remove_from_list,
            context_label="Remove from List",
            selectable=True,
            on_selection_change=self._on_selection_change,
            on_assign_request=self._handle_assign_request if self._is_admin() else None,
        )
        self._tree.grid(row=0, column=0, sticky="nsew")

        # ── Action bar ──────────────────────────────────────────────────────
        self._build_action_bar(parent)

    def _build_action_bar(self, parent):
        self._action_bar = ctk.CTkFrame(
            parent, fg_color=("gray85", "gray20"), corner_radius=8,
        )
        self._action_bar.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        self._action_bar.grid_remove()

        self._action_clear_btn = ctk.CTkButton(
            self._action_bar, text="✕ Clear", width=70, height=28,
            fg_color="gray50", hover_color="gray40",
            command=self._clear_all_checks,
        )
        self._action_clear_btn.pack(side="left", padx=(8, 4), pady=6)

        self._action_count_lbl = ctk.CTkLabel(
            self._action_bar, text="", font=ctk.CTkFont(size=12),
        )
        self._action_count_lbl.pack(side="left", padx=(4, 12))

        self._action_sent_btn = ctk.CTkButton(
            self._action_bar, text="Mark as Sent", width=120, height=28,
            fg_color=("#2E7D32", "#1B5E20"), hover_color=("#256528", "#164A18"),
            command=self._bulk_mark_as_sent,
        )
        self._action_sent_btn.pack(side="left", padx=(0, 6))

        self._action_po_btn = ctk.CTkButton(
            self._action_bar, text="Add to PO", width=100, height=28,
            command=self._bulk_add_to_po,
        )
        self._action_po_btn.pack(side="left", padx=(0, 6))

        self._action_assign_btn = ctk.CTkButton(
            self._action_bar, text="Assign to User", width=130, height=28,
            command=self._bulk_assign,
        )
        self._action_assign_btn.pack(side="left", padx=(0, 6))

    def _on_selection_change(self):
        checked = self._tree.get_checked_orders()
        if not checked:
            self._action_bar.grid_remove()
            return
        n = len(checked)
        self._action_count_lbl.configure(text=f"{n} order{'s' if n != 1 else ''} selected")
        self._action_bar.grid()
        musipos = getattr(self._window, "musipos_client", None)
        self._action_sent_btn.configure(state="normal")
        self._action_po_btn.configure(state="normal" if musipos else "disabled")
        self._action_assign_btn.configure(state="normal" if self._is_admin() else "disabled")

    def _is_admin(self) -> bool:
        cu = getattr(self._window, "current_user", None)
        return bool(cu and cu.get("role") == "admin")

    def _clear_all_checks(self):
        self._tree.clear_checks()

    # ── Entry point ────────────────────────────────────────────────────────────

    def show(self):
        """Called by DailyOpsWindow each time this step is shown."""
        self._neto_orders = list(self._window.search_neto_orders)
        self._ebay_orders = list(self._window.search_ebay_orders)
        self._hidden_order_ids.clear()
        self._ungrouped_order_ids.clear()
        self._collated_groups.clear()
        self._refresh_tables()
        self._update_counts(self._get_visible_orders())

    # ── Data helpers ───────────────────────────────────────────────────────────

    def _get_visible_orders(self) -> list:
        visible = []
        for o in self._neto_orders:
            key = (o.sales_channel or "Neto", o.order_id)
            if key not in self._hidden_order_ids:
                visible.append(o)
        for o in self._ebay_orders:
            if ("eBay", o.order_id) not in self._hidden_order_ids:
                visible.append(o)
        return visible

    def _build_groups(self, orders: list) -> list[dict]:
        from src.order_collator import collate_orders

        neto_orders = [o for o in orders if hasattr(o, "date_placed")]
        ebay_orders = [o for o in orders if not hasattr(o, "date_placed")]
        coll_groups, neto_singles, ebay_singles = collate_orders(
            neto_orders, ebay_orders, self._ungrouped_order_ids
        )
        self._collated_groups.update({g.synthetic_id: g for g in coll_groups})

        _assignments: dict = {}
        _users_cache: dict = {}
        _am = getattr(self._window, "assignment_manager", None)
        _um = getattr(self._window, "user_manager", None)
        if _am:
            try:
                _assignments = _am.get_all_assignments()
            except Exception:
                pass
        if _um:
            try:
                _users_cache = {u["username"]: u for u in _um.get_active_users()}
            except Exception:
                pass

        def _assigned_label(order_id: str) -> str:
            a = _assignments.get(order_id, {})
            u = _users_cache.get(a.get("assigned_to", ""), {})
            return u.get("first_name", "") or a.get("assigned_to", "")

        _filt = getattr(self, "_assign_filter", "All")
        if _filt != "All":
            cu = getattr(self._window, "current_user", None)
            current_username = cu.get("username", "") if cu else ""

            def _passes(order_id: str) -> bool:
                assigned_to = _assignments.get(order_id, {}).get("assigned_to", "")
                if _filt == "Mine":
                    return assigned_to == current_username
                return not assigned_to

            neto_singles = [o for o in neto_singles if _passes(o.order_id)]
            ebay_singles = [o for o in ebay_singles if _passes(o.order_id)]
            coll_groups = []

        result = []
        for g in coll_groups:
            d = self._group_dict_for_collated(g)
            d["assigned"] = ""
            result.append(d)
        for o in neto_singles + ebay_singles:
            d = self._group_dict_for_single(o)
            d["assigned"] = _assigned_label(o.order_id)
            result.append(d)
        return result

    def _group_dict_for_collated(self, g) -> dict:
        first = g.orders[0]
        is_neto = hasattr(first, "date_placed")
        if is_neto:
            ship_name = (
                f"{first.ship_first_name} {first.ship_last_name}".strip()
                or first.customer_name
            )
            order_date = first.date_placed
        else:
            ship_name = first.ship_name or first.buyer_name
            order_date = getattr(first, "creation_date", None)

        all_line_items = []
        total_val = 0.0
        for o in g.orders:
            o_is_neto = hasattr(o, "date_placed")
            total_val += getattr(o, "grand_total" if o_is_neto else "order_total", 0.0) or 0.0
            for li in o.line_items:
                all_line_items.append({
                    "sku": li.sku,
                    "description": li.product_name if o_is_neto else li.title,
                    "qty": str(li.quantity),
                    "is_matched": False,
                })

        notes_parts = []
        for o in g.orders:
            n = (o.notes if hasattr(o, "date_placed") else getattr(o, "buyer_notes", "")) or ""
            if n:
                notes_parts.append(f"[{o.order_id}] {n}")
        order_notes = "  |  ".join(notes_parts)

        total_str = f"${total_val:.2f}" if total_val else "—"
        notes = f"Collated  ·  {total_str}"
        date_str = order_date.strftime("%d/%m") if order_date else "—"

        return {
            "order_id": g.synthetic_id,
            "platform": g.platform,
            "customer": ship_name,
            "date": date_str,
            "shipping": f"{len(g.orders)} orders",
            "notes": notes,
            "status": "",
            "order_notes": order_notes,
            "postage_type": "Mixed",
            "line_items": all_line_items,
        }

    def _group_dict_for_single(self, order) -> dict:
        is_neto = hasattr(order, "date_placed")
        if is_neto:
            platform = order.sales_channel or "Neto"
            ship_name = (
                f"{order.ship_first_name} {order.ship_last_name}".strip()
                or order.customer_name
            )
            shipping = order.shipping_type
            total_val = getattr(order, "grand_total", 0.0) or 0.0
            order_date = order.date_placed
            line_items = [
                {
                    "sku": li.sku,
                    "description": li.product_name,
                    "qty": str(li.quantity),
                    "is_matched": False,
                }
                for li in order.line_items
            ]
            order_notes = order.notes or ""
        else:
            platform = "eBay"
            ship_name = order.ship_name or order.buyer_name
            shipping = order.shipping_type
            total_val = getattr(order, "order_total", 0.0) or 0.0
            order_date = getattr(order, "creation_date", None)
            line_items = [
                {
                    "sku": li.sku,
                    "description": li.title,
                    "qty": str(li.quantity),
                    "is_matched": False,
                }
                for li in order.line_items
            ]
            order_notes = getattr(order, "buyer_notes", "") or ""

        total_str = f"${total_val:.2f}" if total_val else "—"
        date_str = order_date.strftime("%d/%m") if order_date else "—"

        if is_neto:
            status = order.status or ""
        else:
            status = _EBAY_STATUS_DISPLAY.get(order.order_status, order.order_status or "")

        return {
            "order_id": order.order_id,
            "platform": platform,
            "customer": ship_name,
            "date": date_str,
            "shipping": shipping,
            "notes": total_str,
            "status": status,
            "order_notes": order_notes,
            "postage_type": "",
            "line_items": line_items,
        }

    def _find_order_data(self, order_id: str, platform: str):
        if platform.lower() == "ebay":
            for o in self._ebay_orders:
                if o.order_id == order_id:
                    return None, o, []
        else:
            for o in self._neto_orders:
                if o.order_id == order_id:
                    return o, None, []
        return None, None, []

    def _find_order_for_bulk(self, order_id: str, platform: str):
        if platform.lower() == "ebay":
            for o in self._ebay_orders:
                if o.order_id == order_id:
                    return o
        else:
            for o in self._neto_orders:
                if o.order_id == order_id:
                    return o
        return None

    # ── Table population ───────────────────────────────────────────────────────

    def _refresh_tables(self):
        self._collated_groups.clear()
        visible = self._get_visible_orders()
        self._tree.load_groups(self._build_groups(visible))
        self._update_counts(visible)

    def _update_counts(self, visible: list):
        total = len(self._neto_orders) + len(self._ebay_orders)
        hidden = len(self._hidden_order_ids)
        if hidden:
            msg = f"{len(visible)} shown  ·  {hidden} hidden  ·  {total} total"
        else:
            msg = f"{total} order{'s' if total != 1 else ''}"
        self._status_label.configure(text=msg, text_color=("gray50", "gray60"))

    # ── Remove from list ───────────────────────────────────────────────────────

    def _remove_from_list(self, order_id: str, platform: str):
        if order_id in self._collated_groups:
            g = self._collated_groups[order_id]
            for oid in g.order_ids:
                self._hidden_order_ids.add((g.platform, oid))
        else:
            self._hidden_order_ids.add((platform, order_id))
        self._refresh_tables()

    # ── Assignment ─────────────────────────────────────────────────────────────

    def _on_assign_filter_change(self, value: str) -> None:
        self._assign_filter = value
        self._refresh_tables()

    def _handle_assign_request(self, order_ids: list, platform: str,
                               single: bool = False, remove: bool = False) -> None:
        if remove:
            am = getattr(self._window, "assignment_manager", None)
            if am:
                try:
                    am.unassign(order_ids)
                except Exception:
                    pass
            self._refresh_all_orders()
            return
        self._open_assign_popup(order_ids)

    def _bulk_assign(self) -> None:
        checked = self._tree.get_checked_orders()
        if not checked:
            return
        order_ids = [c["order_id"] for c in checked]
        self._open_assign_popup(order_ids)

    def _open_assign_popup(self, order_ids: list) -> None:
        am = getattr(self._window, "assignment_manager", None)
        um = getattr(self._window, "user_manager", None)
        cu = getattr(self._window, "current_user", None)
        if not am or not um or not cu:
            return
        try:
            users = um.get_active_users()
        except Exception:
            return

        popup = ctk.CTkToplevel(self.winfo_toplevel())
        popup.title("Assign Orders")
        popup.geometry("320x180")
        popup.resizable(False, False)
        popup.transient(self.winfo_toplevel())
        popup.grab_set()

        body = ctk.CTkFrame(popup, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        n = len(order_ids)
        ctk.CTkLabel(
            body,
            text=f"Assign {n} order{'s' if n != 1 else ''} to:",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", pady=(0, 10))

        user_options = ["— Unassign —"] + [
            f"{u['first_name']} {u['last_name']}".strip() or u["username"]
            for u in users
        ]
        user_map = {"— Unassign —": ""} | {
            (f"{u['first_name']} {u['last_name']}".strip() or u["username"]): u["username"]
            for u in users
        }
        sel_var = ctk.StringVar(value=user_options[0])
        ctk.CTkComboBox(body, variable=sel_var, values=user_options,
                        state="readonly", height=34).pack(fill="x", pady=(0, 14))

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x")

        def _confirm():
            target_username = user_map.get(sel_var.get(), "")
            try:
                if target_username:
                    am.assign(order_ids, target_username, cu["username"])
                else:
                    am.unassign(order_ids)
            except Exception:
                pass
            self._tree.clear_checks()
            self._refresh_all_orders()
            popup.destroy()

        ctk.CTkButton(btn_row, text="Assign", command=_confirm).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="Cancel", fg_color="transparent",
                      border_width=1, command=popup.destroy).pack(side="left")
        popup.after(50, popup.lift)

    # ── Bulk actions ──────────────────────────────────────────────────────────

    def _bulk_mark_as_sent(self):
        from src.cross_dispatch import sync_ebay_to_neto, sync_neto_to_ebay
        checked = self._tree.get_checked_orders()
        if not checked:
            return
        n = len(checked)
        if not messagebox.askyesno(
            "Mark as Sent",
            f"Mark {n} order{'s' if n != 1 else ''} as sent without tracking numbers?\n\nThis cannot be undone.",
            parent=self,
        ):
            return

        dry_run = self._window.config.app.dry_run
        neto_client = self._window.neto_client
        ebay_client = self._window.ebay_client
        results: list = []
        failed_ebay_syncs: list = []  # (ebay_order_id, error_msg) for dialog

        def _work():
            for c in checked:
                order = self._find_order_for_bulk(c["order_id"], c["platform"])
                if order is None:
                    results.append(f"⚠ {c['order_id']}: not found")
                    continue
                try:
                    if c["platform"].lower() == "ebay":
                        ebay_client.create_shipping_fulfillment(
                            c["order_id"],
                            line_items=order.line_items,
                            tracking_number="",
                            carrier="",
                            dry_run=dry_run,
                        )
                    else:
                        skus = [li.sku for li in order.line_items]
                        neto_client.update_order_status(
                            c["order_id"],
                            new_status="Dispatched",
                            tracking_number="",
                            shipping_method="",
                            line_item_skus=skus,
                            dry_run=dry_run,
                        )
                    label = "[DRY RUN] " if dry_run else ""
                    results.append(f"✓ {label}{c['order_id']}")
                except Exception as exc:
                    results.append(f"✗ {c['order_id']}: {exc}")
                    continue

                # Cross-platform sync (best-effort, non-fatal)
                try:
                    if c["platform"].lower() == "ebay" and neto_client:
                        ok, msg = sync_ebay_to_neto(
                            c["order_id"], neto_client,
                            ebay_created_at=getattr(order, "creation_date", None),
                            dry_run=dry_run,
                        )
                        results.append(f"  {'✓' if ok else '⚠'} Neto sync: {msg}")
                    elif c["platform"].lower() != "ebay" and ebay_client:
                        ebay_id = getattr(order, "purchase_order_number", "")
                        if getattr(order, "sales_channel", "").lower() == "ebay" and ebay_id:
                            ok, msg = sync_neto_to_ebay(
                                order, ebay_client,
                                dry_run=dry_run,
                            )
                            results.append(f"  {'✓' if ok else '⚠'} eBay sync: {msg}")
                            if not ok:
                                failed_ebay_syncs.append((ebay_id, msg))
                except Exception as exc:
                    results.append(f"  ⚠ sync error: {exc}")

            self.after(0, lambda: _done())

        def _done():
            from src.gui.dialogs import show_ebay_sync_failed_dialog
            self._tree.clear_checks()
            messagebox.showinfo("Mark as Sent — Results", "\n".join(results), parent=self)
            if failed_ebay_syncs:
                show_ebay_sync_failed_dialog(self, failed_ebay_syncs)
            self._refresh_all_orders()

        threading.Thread(target=_work, daemon=True).start()

    def _bulk_add_to_po(self):
        musipos_client = getattr(self._window, "musipos_client", None)
        if musipos_client is None:
            messagebox.showinfo("Not configured", "Musipos is not configured.", parent=self)
            return
        checked = self._tree.get_checked_orders()
        if not checked:
            return

        from collections import deque
        items_queue: deque = deque()
        for c in checked:
            order = self._find_order_for_bulk(c["order_id"], c["platform"])
            if order is None:
                continue
            for li in (order.line_items or []):
                items_queue.append((c["order_id"], c["platform"], li))

        if not items_queue:
            messagebox.showinfo("Add to PO", "No line items found in selected orders.", parent=self)
            return

        self._bulk_po_next(items_queue)

    def _bulk_po_next(self, queue):
        if not queue:
            messagebox.showinfo("Add to PO", "All items processed.", parent=self.winfo_toplevel())
            self._refresh_all_orders()
            return

        from collections import deque
        from src.gui.musipos_po_dialog import MusiposPODialog

        order_id, platform, line_item = queue.popleft()
        musipos_client = getattr(self._window, "musipos_client", None)
        if musipos_client is None:
            return

        qty = getattr(line_item, "quantity", 1) or 1
        product_name = (
            getattr(line_item, "product_name", None)
            or getattr(line_item, "title", "")
            or ""
        )

        advanced = [False]

        def _advance():
            if not advanced[0]:
                advanced[0] = True
                self.after(0, lambda: self._bulk_po_next(queue))

        def _on_success(po_result):
            sku = po_result.get("resolved_sku") or line_item.sku
            self._bulk_po_write_note(order_id, platform, sku, line_item)

        def _on_note_only(resolved_sku=None):
            sku = resolved_sku or line_item.sku
            self._bulk_po_write_note(order_id, platform, sku, line_item)

        def _on_cancel():
            pass

        dialog = MusiposPODialog(
            self.winfo_toplevel(),
            neto_sku=line_item.sku,
            product_name=product_name,
            order_qty=qty,
            musipos_client=musipos_client,
            suppliers_config=self._window.config.suppliers,
            dry_run=self._window.config.app.dry_run,
            on_success=_on_success,
            on_note_only=_on_note_only,
            on_cancel=_on_cancel,
        )
        dialog.protocol("WM_DELETE_WINDOW", lambda: (dialog.destroy(), _advance()))
        dialog.bind("<Destroy>", lambda e: _advance() if e.widget is dialog else None)

    def _bulk_po_write_note(self, order_id: str, platform: str, sku: str, line_item=None):
        from datetime import date
        dry_run = self._window.config.app.dry_run
        note_text = f"{sku} on PO"

        if platform.lower() == "ebay":
            ebay_client = self._window.ebay_client
            if not ebay_client or line_item is None:
                return
            item_id = getattr(line_item, "legacy_item_id", "")
            txn_id = getattr(line_item, "legacy_transaction_id", "")
            if not item_id:
                return

            def _ebay_work():
                try:
                    ebay_client.set_private_notes(
                        item_id=item_id,
                        transaction_id=txn_id,
                        note_text=note_text[:255],
                        dry_run=dry_run,
                    )
                except Exception:
                    pass

            threading.Thread(target=_ebay_work, daemon=True).start()
            return

        neto_client = self._window.neto_client
        if not neto_client:
            return
        dated_note = f"[{date.today().strftime('%d/%m/%Y')}] {note_text}"

        def _neto_work():
            try:
                neto_client.add_sticky_note(order_id, title="Item Status",
                                            description=dated_note, dry_run=dry_run)
            except Exception:
                pass

        threading.Thread(target=_neto_work, daemon=True).start()

    # ── Refresh (re-runs full search) ──────────────────────────────────────────

    def _refresh_all_orders(self):
        if not self._search_options:
            return
        self._refresh_btn.configure(state="disabled")
        self._error_label.configure(text="Refreshing…")
        threading.Thread(
            target=self._refresh_worker,
            args=(self._search_options,),
            daemon=True,
        ).start()

    def _refresh_worker(self, options: dict):
        fast_path = options.get("order_id_fast_path")
        if fast_path:
            self._fast_path_refresh(options, fast_path)
            return

        # ── Normal path: run Neto + eBay in parallel ───────────────────────
        neto_orders: list = []
        ebay_orders: list = []
        neto_error: list[str] = []
        ebay_warning: list[str] = []

        def _fetch_neto():
            neto_statuses = options.get("neto_statuses", [])
            any_neto_platform = any(options.get("platforms", {}).values())
            if not (neto_statuses and any_neto_platform):
                return
            try:
                terms = options.get("search_terms", {})
                raw = self._window.neto_client.search_orders(
                    options["date_from"], options["date_to"],
                    statuses=neto_statuses,
                    sku=terms.get("sku", "").strip(),
                    title=terms.get("title", "").strip(),
                    customer_name=terms.get("name", "").strip(),
                )
                neto_orders.extend(_filter_neto_by_channel(raw, options))
            except (NetoAPIError, Exception) as exc:
                neto_error.append(str(exc))

        def _fetch_ebay():
            ebay_statuses = options.get("ebay_fulfillment_statuses", [])
            if not (options.get("ebay_direct", True) and ebay_statuses):
                return
            try:
                results = self._window.ebay_client.search_orders(
                    options["date_from"], options["date_to"],
                    fulfillment_statuses=ebay_statuses,
                    enrich_notes=False,
                )
                ebay_orders.extend(results)
            except (EbayAuthError, EbayAPIError, Exception) as exc:
                ebay_warning.append(str(exc))
                log.warning("Refresh eBay search failed (continuing with Neto): %s", exc)

        neto_thread = threading.Thread(target=_fetch_neto, daemon=True)
        ebay_thread = threading.Thread(target=_fetch_ebay, daemon=True)
        neto_thread.start()
        ebay_thread.start()
        neto_thread.join()
        ebay_thread.join()

        if neto_error:
            self.after(0, lambda m=neto_error[0]: self._on_refresh_error(f"Neto: {m}"))
            return

        terms = options.get("search_terms", {})
        filtered_neto, filtered_ebay = _apply_text_filters(neto_orders, ebay_orders, terms)
        warning = ebay_warning[0] if ebay_warning else ""
        self.after(
            0,
            lambda n=filtered_neto, e=filtered_ebay, w=warning:
                self._on_refresh_done(n, e, w),
        )

    def _fast_path_refresh(self, options: dict, fast_path: dict):
        """Targeted single-order refresh when the original search was by exact order ID."""
        platform = fast_path["platform"]
        order_id = fast_path["order_id"]
        neto_orders: list = []
        ebay_orders: list = []
        error_msg: str | None = None
        ebay_warning: str = ""

        if platform == "neto":
            try:
                neto_orders = self._window.neto_client.get_order_by_exact_id(order_id)
            except (NetoAPIError, Exception) as exc:
                error_msg = f"Neto: {exc}"
        else:
            try:
                ebay_orders = self._window.ebay_client.get_order_by_exact_id(order_id)
            except (EbayAuthError, EbayAPIError, Exception) as exc:
                ebay_warning = str(exc)

        if error_msg:
            self.after(0, lambda msg=error_msg: self._on_refresh_error(f"Refresh failed: {msg}"))
            return

        terms = {k: v for k, v in options.get("search_terms", {}).items() if k != "order_number"}
        filtered_neto, filtered_ebay = _apply_text_filters(neto_orders, ebay_orders, terms)
        self.after(
            0,
            lambda n=filtered_neto, e=filtered_ebay, w=ebay_warning:
                self._on_refresh_done(n, e, w),
        )

    def _on_refresh_done(self, neto_orders: list, ebay_orders: list, ebay_warning: str):
        self._neto_orders = neto_orders
        self._ebay_orders = ebay_orders
        self._assign_filter = "All"
        self._assign_seg.set("All")
        self._refresh_tables()
        self._refresh_btn.configure(state="normal")
        msg = f"Refreshed {datetime.now().strftime('%H:%M')}"
        if ebay_warning:
            msg += f"  ⚠ eBay: {ebay_warning}"
            self._error_label.configure(text=f"eBay: {ebay_warning}")
        else:
            self._error_label.configure(text="")
        self._status_label.configure(text=msg, text_color=("gray50", "gray60"))

    def _on_refresh_error(self, msg: str):
        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text=f"Refresh failed: {msg}")

    # ── Order Detail navigation ────────────────────────────────────────────────

    def _open_detail_view(self, order_id: str, platform: str):
        if order_id in self._collated_groups:
            self._open_collated_view(self._collated_groups[order_id])
            return

        from src.gui.order_detail_view import OrderDetailView

        self._last_clicked_order_id = order_id
        self._last_clicked_platform = platform

        neto_order, ebay_order, _ = self._find_order_data(order_id, platform)
        if neto_order is None and ebay_order is None:
            self._error_label.configure(text=f"Order {order_id} not found")
            return

        um = getattr(self._window, "user_manager", None)
        cu = getattr(self._window, "current_user", None)
        if um and cu:
            other = um.get_processing_user(order_id)
            if other and other != f"{cu.get('first_name','')} {cu.get('last_name','')}".strip():
                if not messagebox.askyesno(
                    "Order In Progress",
                    f"{other} is currently processing order #{order_id}.\n\nTake over?",
                    parent=self.winfo_toplevel(),
                ):
                    return
            um.set_processing_order(cu["username"], order_id)
            app = getattr(self._window, "master", None)
            if app and hasattr(app, "_processing_order_id"):
                app._processing_order_id = order_id
                app._order_close_callback = self._clear_processing_flag

        if self._detail_frame is not None:
            self._detail_frame.destroy()

        book_freight_cb = None
        if self._window.config.shipping is not None:
            book_freight_cb = self._open_freight_view

        self._detail_frame = OrderDetailView(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            matched_skus=[],
            neto_client=self._window.neto_client,
            ebay_client=self._window.ebay_client,
            dry_run=self._window.config.app.dry_run,
            on_back=self._close_detail_view,
            on_fulfilled=self._on_fulfilled,
            on_move_to_unmatched=lambda: self._remove_from_list(order_id, platform),
            move_to_unmatched_label="Remove from List",
            on_book_freight=book_freight_cb,
            sku_alias_manager=self._window.sku_alias_manager,
            suppliers=self._window.config.suppliers,
            musipos_client=getattr(self._window, "musipos_client", None),
            variation_manager=getattr(self._window, "ebay_variation_manager", None),
        )
        self._detail_frame.grid(row=0, column=0, sticky="nsew")
        self._detail_frame.tkraise()

    def _clear_processing_flag(self):
        um = getattr(self._window, "user_manager", None)
        cu = getattr(self._window, "current_user", None)
        if um and cu:
            try:
                um.clear_processing_order(cu["username"])
            except Exception:
                pass
        app = getattr(self._window, "master", None)
        if app and hasattr(app, "_processing_order_id"):
            app._processing_order_id = None
            app._order_close_callback = None

    def _close_detail_view(self):
        self._clear_processing_flag()
        if self._detail_frame is not None:
            self._detail_frame.destroy()
            self._detail_frame = None
        self._list_frame.tkraise()
        if self._last_clicked_order_id:
            self._tree.scroll_to(self._last_clicked_order_id)
        self._refresh_all_orders()

    def _on_fulfilled(self):
        if self._last_clicked_order_id:
            am = getattr(self._window, "assignment_manager", None)
            if am:
                try:
                    am.clear_on_dispatch(self._last_clicked_order_id)
                except Exception:
                    pass
        self._close_detail_view()

    # ── Collated detail view ───────────────────────────────────────────────────

    def _open_collated_view(self, group):
        from src.gui.daily_ops.collated_detail_view import CollatedDetailView

        if self._collated_frame is not None:
            self._collated_frame.destroy()

        self._collated_frame = CollatedDetailView(
            self,
            group=group,
            window=self._window,
            dry_run=self._window.config.app.dry_run,
            on_back=self._close_collated_view,
            on_ungroup=self._ungroup_orders,
            on_book_freight=self._on_collated_book_freight,
        )
        self._collated_frame.grid(row=0, column=0, sticky="nsew")
        self._collated_frame.tkraise()

    def _close_collated_view(self):
        if self._collated_frame is not None:
            self._collated_frame.destroy()
            self._collated_frame = None
        self._list_frame.tkraise()
        self._refresh_all_orders()

    def _ungroup_orders(self, order_ids: list):
        self._ungrouped_order_ids.update(order_ids)
        self._refresh_tables()

    def _on_collated_book_freight(self, order_id: str, platform: str, for_all: bool = False):
        self._book_freight_for_all = for_all
        self._open_freight_view(order_id, platform)

    # ── Freight booking ────────────────────────────────────────────────────────

    def _open_freight_view(self, order_id: str, platform: str):
        from src.gui.freight_booking_view import FreightBookingView

        self._last_clicked_order_id = order_id
        self._last_clicked_platform = platform

        neto_order, ebay_order, _ = self._find_order_data(order_id, platform)

        if self._freight_frame is not None:
            self._freight_frame.destroy()

        self._freight_frame = FreightBookingView(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            neto_client=self._window.neto_client,
            ebay_client=self._window.ebay_client,
            shipping_config=self._window.config.shipping,
            dry_run=self._window.config.app.dry_run,
            on_back=self._close_freight_view,
            on_courier_selected=lambda name, tracking="": self._on_courier_selected(name, tracking),
        )
        self._freight_frame.grid(row=0, column=0, sticky="nsew")
        self._freight_frame.tkraise()

    def _close_freight_view(self):
        if self._freight_frame is not None:
            self._freight_frame.destroy()
            self._freight_frame = None
        if self._collated_frame is not None:
            self._collated_frame.tkraise()
        elif self._detail_frame is not None:
            self._detail_frame.tkraise()

    def _on_courier_selected(self, courier_name: str, tracking_number: str = ""):
        self._close_freight_view()
        if self._book_freight_for_all and self._collated_frame is not None:
            if tracking_number:
                self._collated_frame.set_tracking_all(tracking_number, courier_name)
                self._collated_frame._mark_all_as_sent()
            self._book_freight_for_all = False
            return
        self._book_freight_for_all = False
        if self._collated_frame is not None:
            if tracking_number and self._last_clicked_order_id:
                self._collated_frame.set_tracking_for(
                    self._last_clicked_order_id, tracking_number, courier_name
                )
            return
        if self._detail_frame is not None:
            self._detail_frame.set_tracking(tracking=tracking_number, carrier=courier_name)
            if tracking_number:
                self._detail_frame._mark_as_sent()

    # ── Cancel / Reprint shipment dialog ──────────────────────────────────────

    def _open_cancel_dialog(self):
        import tkinter.ttk as ttk
        from pathlib import Path

        shipping = self._window.config.shipping
        if shipping is None:
            messagebox.showerror("Not configured", "Shipping is not configured.", parent=self)
            return

        bookings_dir = shipping.bookings_dir
        if not bookings_dir:
            messagebox.showerror("Not configured", "Bookings directory is not configured.", parent=self)
            return

        from src.shipping.booking_ledger import get_all_bookings, mark_cancelled
        from src.shipping.couriers.allied import AlliedCourier
        from src.shipping.couriers.aramex import AramexCourier
        from src.shipping.couriers.auspost import AusPostCourier
        from src.shipping.couriers.bonds import BondsCourier
        from src.shipping.couriers.dai_post import DaiPostCourier

        all_bookings = get_all_bookings(bookings_dir, days=1)

        courier_registry = {
            "auspost": AusPostCourier,
            "aramex": AramexCourier,
            "bonds": BondsCourier,
            "allied": AlliedCourier,
            "dai_post": DaiPostCourier,
        }
        couriers_by_code = {}
        for code, cls in courier_registry.items():
            cfg = shipping.couriers.get(code, {})
            if cfg.get("enabled", False):
                couriers_by_code[code] = cls(cfg)

        _sort_col = ["date"]
        _sort_asc = [False]

        win = tk.Toplevel(self)
        win.title("Manage Shipments")
        win.resizable(True, False)
        win.grab_set()

        tk.Label(
            win, text="Select a booking to cancel or reprint its label:",
            font=("Segoe UI", 11, "bold"),
        ).pack(padx=16, pady=(12, 6), anchor="w")

        tree_frame = tk.Frame(win)
        tree_frame.pack(fill="x", padx=16, pady=(0, 8))

        columns = ("date", "time", "courier", "order", "recipient", "tracking")
        col_labels = {
            "date": "Date", "time": "Time", "courier": "Courier",
            "order": "Order", "recipient": "Recipient", "tracking": "Tracking #",
        }
        col_widths = {
            "date": 90, "time": 55, "courier": 120, "order": 90,
            "recipient": 140, "tracking": 160,
        }
        col_stretch = {
            "date": False, "time": False, "courier": False,
            "order": False, "recipient": True, "tracking": True,
        }

        tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            height=min(max(len(all_bookings), 1), 12),
            selectmode="browse",
        )
        for col in columns:
            tree.heading(col, text=col_labels[col], command=lambda c=col: _sort_by(c))
            tree.column(col, width=col_widths[col], stretch=col_stretch[col])

        _iid_to_booking: dict[str, dict] = {}

        def _col_val(b: dict, col: str) -> str:
            if col == "time":
                t = b.get("booked_at", "")
                return t.split("T")[1][:5] if "T" in t else ""
            key_map = {
                "date": "date", "courier": "courier_name",
                "order": "order_id", "recipient": "recipient",
                "tracking": "tracking_number",
            }
            return b.get(key_map.get(col, col), "")

        def _populate():
            tree.delete(*tree.get_children())
            _iid_to_booking.clear()
            if not all_bookings:
                tree.insert("", "end", values=("", "", "No bookings found", "", "", ""))
                return
            col, asc = _sort_col[0], _sort_asc[0]
            sort_key = (
                (lambda b: b.get("booked_at", ""))
                if col in ("date", "time")
                else (lambda b: _col_val(b, col))
            )
            for b in sorted(all_bookings, key=sort_key, reverse=not asc):
                iid = tree.insert("", "end", values=tuple(_col_val(b, c) for c in columns))
                _iid_to_booking[iid] = b

        def _sort_by(col: str):
            if _sort_col[0] == col:
                _sort_asc[0] = not _sort_asc[0]
            else:
                _sort_col[0] = col
                _sort_asc[0] = True
            for c in columns:
                arrow = (" ▲" if _sort_asc[0] else " ▼") if c == _sort_col[0] else ""
                tree.heading(c, text=col_labels[c] + arrow)
            _populate()

        tree.heading("date", text="Date ▼")
        _populate()

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tree_frame.grid_columnconfigure(0, weight=1)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=(4, 4))

        cancel_btn = tk.Button(
            btn_frame, text="Cancel Shipment", font=("Segoe UI", 10),
            bg="#b22222", fg="white", activebackground="#8b0000",
            width=20, state="disabled", command=lambda: _confirm_cancel(),
        )
        cancel_btn.pack(side="left", padx=(0, 8))

        reprint_btn = tk.Button(
            btn_frame, text="Reprint Label", font=("Segoe UI", 10),
            bg="#1a6b1a", fg="white", activebackground="#0f4a0f",
            width=16, state="disabled", command=lambda: _reprint_label(),
        )
        reprint_btn.pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame, text="Close", font=("Segoe UI", 10),
            width=10, command=win.destroy,
        ).pack(side="left")

        status_lbl = tk.Label(win, text="", font=("Segoe UI", 10), wraplength=560, fg="gray40")
        status_lbl.pack(padx=16, pady=(4, 12))

        def _on_select(_event=None):
            state = "normal" if tree.selection() and all_bookings else "disabled"
            cancel_btn.configure(state=state)
            reprint_btn.configure(state=state)

        tree.bind("<<TreeviewSelect>>", _on_select)
        tree.bind("<Double-1>", lambda _e: _reprint_label())

        def _reprint_label():
            sel = tree.selection()
            if not sel:
                return
            booking = _iid_to_booking.get(sel[0])
            if not booking:
                return
            order_id = booking.get("order_id", "")
            booking_date = booking.get("date", "")
            courier_code = booking.get("print_courier_code") or booking.get("courier_code", "")
            label_path = Path(bookings_dir) / "Labels" / booking_date / f"{order_id}.pdf"
            if not label_path.exists():
                status_lbl.configure(text=f"Label not found:\n{label_path}", fg="red")
                return
            pdf_bytes = label_path.read_bytes()
            reprint_btn.configure(state="disabled")
            status_lbl.configure(text="Printing…", fg="gray40")
            win.update_idletasks()

            def _run():
                from src.shipping.label_printer import print_label
                err = print_label(pdf_bytes, courier_code=courier_code)
                win.after(0, lambda: _on_print_done(err, order_id))

            def _on_print_done(err: str, oid: str):
                reprint_btn.configure(state="normal")
                if err:
                    status_lbl.configure(text=f"Print failed: {err}", fg="red")
                else:
                    status_lbl.configure(text=f"Label for {oid} sent to printer.", fg="green")

            threading.Thread(target=_run, daemon=True).start()

        def _confirm_cancel():
            sel = tree.selection()
            if not sel:
                return
            iid = sel[0]
            booking = _iid_to_booking.get(iid)
            if not booking:
                return

            courier_code = booking.get("courier_code", "")
            courier_name = booking.get("courier_name", "")
            tracking = booking.get("tracking_number", "")
            booking_date = booking.get("date", "")

            courier = couriers_by_code.get(courier_code)
            if courier is None:
                status_lbl.configure(
                    text=f"Courier '{courier_name}' is not enabled — cannot cancel via API.",
                    fg="red",
                )
                return

            confirmed = messagebox.askyesno(
                "Confirm Cancellation",
                f"Cancel {courier_name} shipment?\n\n"
                f"Order:    {booking.get('order_id', '')}\n"
                f"Tracking: {tracking}\n"
                f"Date:     {booking_date}\n\n"
                "This cannot be undone.",
                parent=win,
            )
            if not confirmed:
                return

            cancel_btn.configure(state="disabled")
            reprint_btn.configure(state="disabled")
            status_lbl.configure(text="Cancelling…", fg="gray40")
            win.update_idletasks()

            def _run():
                try:
                    result = courier.cancel_shipment(
                        tracking_number=tracking,
                        dry_run=self._window.config.app.dry_run,
                    )
                    win.after(0, lambda r=result: _on_cancel_done(r, iid))
                except Exception as exc:
                    win.after(0, lambda m=str(exc): _on_cancel_error(m))

            def _on_cancel_done(result: dict, cancelled_iid: str):
                if result.get("success"):
                    mark_cancelled(bookings_dir, tracking)
                    tree.delete(cancelled_iid)
                    _iid_to_booking.pop(cancelled_iid, None)
                    status_lbl.configure(
                        text=f"Shipment {tracking} cancelled successfully.", fg="green"
                    )
                else:
                    msg = result.get("message", "Unknown error")
                    status_lbl.configure(text=f"Cancellation failed: {msg}", fg="red")
                cancel_btn.configure(state="normal")
                reprint_btn.configure(state="normal")

            def _on_cancel_error(msg: str):
                status_lbl.configure(text=f"Error: {msg}", fg="red")
                cancel_btn.configure(state="normal")
                reprint_btn.configure(state="normal")

            threading.Thread(target=_run, daemon=True).start()
