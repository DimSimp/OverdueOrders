from __future__ import annotations

from typing import Callable, Optional

import customtkinter as ctk


class CustomerInfoTab(ctk.CTkFrame):
    """Read-only display of all customer fields with an Edit button."""

    def __init__(self, parent, on_edit: Optional[Callable] = None, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._on_edit = on_edit
        self._customer: Optional[dict] = None
        self._layout_cols = 2
        self._value_wrap = 260

        self._build_ui()

    def _build_ui(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 4))

        self._lbl_name = ctk.CTkLabel(
            header,
            text="",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        )
        self._lbl_name.pack(side="left")

        # Active/inactive status badge is intentionally hidden for now.
        # We have not used customer status operationally, but the original
        # badge code is left here in case the feature returns later.
        #
        # self._lbl_badge = ctk.CTkLabel(
        #     header,
        #     text="",
        #     font=ctk.CTkFont(size=12),
        #     width=70,
        #     corner_radius=6,
        # )
        # self._lbl_badge.pack(side="left", padx=(10, 0))

        if self._on_edit:
            ctk.CTkButton(
                header,
                text="Edit",
                width=70,
                height=28,
                font=ctk.CTkFont(size=12),
                command=self._do_edit,
            ).pack(side="right")

        self._scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._scroll.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self._grid_frame = ctk.CTkFrame(self._scroll, fg_color="transparent")
        self._grid_frame.pack(fill="x")

        self.bind("<Configure>", self._on_resize)

    def load_customer(self, customer: dict):
        self._customer = customer
        self._refresh()

    def clear(self):
        self._customer = None
        self._lbl_name.configure(text="")
        # if hasattr(self, "_lbl_badge"):
        #     self._lbl_badge.configure(text="", fg_color="transparent")
        for widget in self._grid_frame.winfo_children():
            widget.destroy()

    def _refresh(self):
        c = self._customer or {}

        first = c.get("first_name") or ""
        surname = c.get("surname") or ""
        self._lbl_name.configure(text=f"{first} {surname}".strip() or "(No name)")

        # active = c.get("active", True)
        # if hasattr(self, "_lbl_badge"):
        #     self._lbl_badge.configure(
        #         text="Active" if active else "Inactive",
        #         fg_color=("#1a7a4a", "#1a7a4a") if active else ("#8b1a1a", "#8b1a1a"),
        #         text_color="white",
        #     )

        for widget in self._grid_frame.winfo_children():
            widget.destroy()

        addr_parts = [
            c.get("address_1"),
            c.get("address_2"),
            c.get("city"),
            c.get("state"),
            c.get("postcode"),
            c.get("country"),
        ]
        address = ", ".join(p for p in addr_parts if p)

        ship_parts = [
            c.get("ship_address_1"),
            c.get("ship_address_2"),
            c.get("ship_city"),
            c.get("ship_state"),
            c.get("ship_postcode"),
            c.get("ship_country"),
        ]
        shipping_address = ", ".join(p for p in ship_parts if p)

        terms_val = c.get("terms_days")
        terms_str = f"{terms_val} days" if terms_val else "-"

        credit = c.get("credit_limit")
        credit_str = f"${credit:,.2f}" if credit is not None else "-"
        discount_profile = c.get("discount_profile") or "-"

        field_groups = [
            [
                ("Customer ID", _fmt(c.get("customer_code") or c.get("customer_id"))),
                ("First Name", _fmt(c.get("first_name"))),
                ("Surname", _fmt(c.get("surname"))),
                ("Business", _fmt(c.get("business"))),
                ("Mobile", _fmt(c.get("mobile"))),
                ("Phone", _fmt(c.get("phone_1"))),
                ("Email", _fmt(c.get("email"))),
            ],
            [
                ("Invoice Address", address or "-"),
                ("Shipping Address", shipping_address or "-"),
                ("Profile Discount", discount_profile),
                ("ABN", _fmt(c.get("abn"))),
                ("Credit Limit", credit_str),
                ("Stop Credit", "Yes" if c.get("stop_credit") else "No"),
                ("Created", _fmt_date(c.get("created_at"))),
            ],
            [
                ("Payment Terms", terms_str),
                ("Is Local", "Yes" if c.get("is_local") else "No"),
                ("Private Comment", _fmt(c.get("private_comment"))),
                ("Statement Comment", _fmt(c.get("statement_comment"))),
                ("Newsletter", "Yes" if c.get("newsletter_opt_in") else "No"),
                ("Musipos ID", _fmt(c.get("musipos_account_code"))),
                ("Tax Exemption #", _fmt(c.get("tax_exemption_number"))),
            ],
        ]

        cols = self._layout_cols
        for pair_idx in range(max(cols, 3) * 2):
            self._grid_frame.grid_columnconfigure(pair_idx, weight=0)
        for value_col in range(cols):
            self._grid_frame.grid_columnconfigure(value_col * 2 + 1, weight=1)

        if cols >= 3:
            fields_to_render: list[tuple[int, int, str, str]] = []
            for col_group, group in enumerate(field_groups):
                for grid_row, (label, value) in enumerate(group):
                    fields_to_render.append((col_group, grid_row, label, value))
        else:
            fields = [field for group in field_groups for field in group]
            fields_to_render = []
            for idx, (label, value) in enumerate(fields):
                col_group = idx % cols
                grid_row = idx // cols
                fields_to_render.append((col_group, grid_row, label, value))

        for col_group, grid_row, label, value in fields_to_render:
            col = col_group * 2

            ctk.CTkLabel(
                self._grid_frame,
                text=label + ":",
                font=ctk.CTkFont(size=12),
                text_color=("gray50", "gray60"),
                anchor="e",
                width=128,
            ).grid(row=grid_row, column=col, sticky="ne", padx=(8, 4), pady=4)

            ctk.CTkLabel(
                self._grid_frame,
                text=value,
                font=ctk.CTkFont(size=12),
                anchor="w",
                justify="left",
                wraplength=self._value_wrap,
            ).grid(row=grid_row, column=col + 1, sticky="nw", padx=(0, 20), pady=4)

    def _do_edit(self):
        if self._on_edit and self._customer:
            self._on_edit(self._customer)

    def _on_resize(self, _event=None):
        width = max(self.winfo_width(), 1)
        if width >= 1100:
            cols = 3
            wrap = 180
        elif width >= 760:
            cols = 2
            wrap = 260
        else:
            cols = 1
            wrap = 520

        if cols == self._layout_cols and wrap == self._value_wrap:
            return

        self._layout_cols = cols
        self._value_wrap = wrap
        if self._customer:
            self._refresh()


def _fmt(val) -> str:
    if val is None or val == "":
        return "-"
    return str(val)


def _fmt_date(val) -> str:
    if not val:
        return "-"
    try:
        return str(val)[:10]
    except Exception:
        return str(val)
