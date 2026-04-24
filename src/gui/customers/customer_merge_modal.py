from __future__ import annotations

import threading
from tkinter import messagebox
from typing import Callable, Optional

import customtkinter as ctk


_FIELD_SECTIONS = [
    (
        "General Details",
        [
            ("first_name", "First Name"),
            ("surname", "Surname"),
            ("business", "Business / School"),
            ("mobile", "Mobile"),
            ("phone_1", "Phone"),
            ("email", "Email"),
            ("fax", "Fax"),
            ("website", "Website"),
        ],
    ),
    (
        "Invoice Address",
        [
            ("address_1", "Address Line 1"),
            ("address_2", "Address Line 2"),
            ("city", "City"),
            ("state", "State"),
            ("postcode", "Postcode"),
            ("country", "Country"),
        ],
    ),
    (
        "Shipping Address",
        [
            ("ship_address_1", "Address Line 1"),
            ("ship_address_2", "Address Line 2"),
            ("ship_city", "City"),
            ("ship_state", "State"),
            ("ship_postcode", "Postcode"),
            ("ship_country", "Country"),
            ("ship_same_as_invoice", "Same as Invoice"),
        ],
    ),
    (
        "Account Details",
        [
            ("discount_profile", "Profile Discount"),
            ("terms_days", "Payment Terms"),
            ("credit_limit", "Credit Limit"),
            ("stop_credit", "Stop Credit"),
            ("is_local", "Is Local"),
            ("newsletter_opt_in", "Newsletter"),
            ("abn", "ABN"),
            ("tax_exemption_number", "Tax Exemption #"),
        ],
    ),
    (
        "Comments and System",
        [
            ("private_comment", "Private Comment"),
            ("statement_comment", "Statement Comment"),
            ("musipos_account_code", "Musipos ID"),
            ("musipos_barcode_ref", "Musipos Barcode Ref"),
        ],
    ),
]

_MULTILINE_FIELDS = {
    "private_comment",
    "statement_comment",
}


class CustomerMergeModal(ctk.CTkToplevel):
    """Choose per-field values when merging two customer profiles."""

    def __init__(
        self,
        master,
        customer_a: dict,
        customer_b: dict,
        merged_by: str | None = None,
        on_merged: Optional[Callable[[dict], None]] = None,
    ):
        super().__init__(master)
        self._customer_a = customer_a
        self._customer_b = customer_b
        self._merged_by = merged_by
        self._on_merged = on_merged
        self._field_vars: dict[str, ctk.StringVar] = {}

        self.title(
            f"Merge Customers {self._customer_ref(customer_a)} and {self._customer_ref(customer_b)}"
        )
        self.geometry("1260x860")
        self.minsize(1080, 760)
        self.resizable(True, True)

        self.transient(master)
        self.grab_set()

        self._build_ui()

    def _build_ui(self):
        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)

        header = ctk.CTkFrame(body, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 10))
        header.grid_columnconfigure(0, weight=1, uniform="merge_header")
        header.grid_columnconfigure(1, weight=1, uniform="merge_header")

        self._build_customer_summary(header, 0, "Profile A", self._customer_a, padx=(0, 8))
        self._build_customer_summary(header, 1, "Profile B", self._customer_b, padx=(8, 0))

        ctk.CTkLabel(
            body,
            text=(
                "Choose which value to keep for each field. "
                "Only one profile can supply each final value."
            ),
            font=ctk.CTkFont(size=12),
            text_color=("gray45", "gray65"),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 8))

        for title, fields in _FIELD_SECTIONS:
            self._build_section(body, title, fields)

        footer = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), height=54)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        self._lbl_error = ctk.CTkLabel(
            footer,
            text="",
            text_color=("red", "#e74c3c"),
            font=ctk.CTkFont(size=11),
        )
        self._lbl_error.pack(side="left", padx=16)

        ctk.CTkButton(
            footer,
            text="Cancel",
            width=90,
            height=32,
            fg_color="transparent",
            border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self.destroy,
        ).pack(side="right", padx=(0, 12), pady=10)

        self._btn_merge = ctk.CTkButton(
            footer,
            text="Merge Customers",
            width=150,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._on_merge,
        )
        self._btn_merge.pack(side="right", padx=(0, 8), pady=10)

    def _build_customer_summary(self, parent, column: int, label: str, customer: dict, padx=(0, 0)):
        card = ctk.CTkFrame(
            parent,
            fg_color=("gray92", "gray17"),
            corner_radius=10,
            border_width=1,
            border_color=("gray78", "gray28"),
        )
        card.grid(row=0, column=column, sticky="nsew", padx=padx)

        ctk.CTkLabel(
            card,
            text=label,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=14, pady=(12, 4))

        ctk.CTkLabel(
            card,
            text=self._customer_heading(customer),
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=14)

        ctk.CTkLabel(
            card,
            text=f"Customer ID: {self._customer_ref(customer)}",
            font=ctk.CTkFont(size=11),
            text_color=("gray45", "gray65"),
            anchor="w",
        ).pack(fill="x", padx=14, pady=(2, 12))

    def _build_section(self, parent, title: str, fields: list[tuple[str, str]]):
        card = ctk.CTkFrame(
            parent,
            fg_color=("gray92", "gray17"),
            corner_radius=10,
            border_width=1,
            border_color=("gray78", "gray28"),
        )
        card.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 8))

        card.grid_columnconfigure(0, weight=1, uniform=f"{title}_panels")
        card.grid_columnconfigure(1, weight=1, uniform=f"{title}_panels")

        panel_a = self._section_panel(card, "Profile A")
        panel_a.grid(row=1, column=0, sticky="nsew", padx=(14, 8), pady=(0, 14))

        panel_b = self._section_panel(card, "Profile B")
        panel_b.grid(row=1, column=1, sticky="nsew", padx=(8, 14), pady=(0, 14))

        for idx, (field_key, label) in enumerate(fields):
            value_a = self._customer_a.get(field_key)
            value_b = self._customer_b.get(field_key)

            var = ctk.StringVar(value=self._default_choice(value_a, value_b))
            self._field_vars[field_key] = var

            pady = (2, 8) if field_key in _MULTILINE_FIELDS else 4

            self._build_panel_row(
                panel_a,
                row=idx,
                choice_var=var,
                choice_value="A",
                choice_text="A",
                label=label,
                value=self._format_value(field_key, value_a),
                pady=pady,
            )
            self._build_panel_row(
                panel_b,
                row=idx,
                choice_var=var,
                choice_value="B",
                choice_text="B",
                label=label,
                value=self._format_value(field_key, value_b),
                pady=pady,
            )

    def _section_panel(self, parent, title: str):
        panel = ctk.CTkFrame(
            parent,
            fg_color=("gray95", "gray15"),
            corner_radius=8,
            border_width=1,
            border_color=("gray80", "gray25"),
        )
        panel.grid_columnconfigure(0, weight=0)
        panel.grid_columnconfigure(1, weight=0)
        panel.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(
            panel,
            text=title,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray45", "gray65"),
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(10, 8))
        return panel

    def _build_panel_row(
        self,
        panel,
        *,
        row: int,
        choice_var,
        choice_value: str,
        choice_text: str,
        label: str,
        value: str,
        pady,
    ):
        grid_row = row + 1
        ctk.CTkRadioButton(
            panel,
            text=choice_text,
            variable=choice_var,
            value=choice_value,
            width=44,
            font=ctk.CTkFont(size=11),
        ).grid(row=grid_row, column=0, sticky="nw", padx=(12, 6), pady=pady)

        ctk.CTkLabel(
            panel,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color=("gray45", "gray65"),
            anchor="w",
            width=128,
        ).grid(row=grid_row, column=1, sticky="nw", padx=(0, 8), pady=pady)

        ctk.CTkLabel(
            panel,
            text=value,
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left",
            wraplength=220,
        ).grid(row=grid_row, column=2, sticky="nw", padx=(0, 12), pady=pady)

    def _on_merge(self):
        self._lbl_error.configure(text="")

        selected = self._selected_values()
        if not (selected.get("first_name") or "").strip():
            self._lbl_error.configure(text="A First Name must be selected for the merged profile.")
            return

        confirmed = messagebox.askyesno(
            "Finalize Customer Merge",
            (
                "This merge will create a new customer profile, move the sales history, "
                "and delete both original profiles.\n\nThis action is final. Continue?"
            ),
            parent=self,
        )
        if not confirmed:
            return

        self._btn_merge.configure(state="disabled", text="Merging...")

        def _thread():
            from src.customers.customer_client import merge_customers

            try:
                result = merge_customers(
                    self._customer_a["id"],
                    self._customer_b["id"],
                    selected,
                    merged_by=self._merged_by,
                )
                self.after(0, lambda: self._on_success(result))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._on_error(err))

        threading.Thread(target=_thread, daemon=True).start()

    def _selected_values(self) -> dict:
        selected: dict = {}
        for field_key, var in self._field_vars.items():
            if var.get() == "B":
                selected[field_key] = self._customer_b.get(field_key)
            else:
                selected[field_key] = self._customer_a.get(field_key)
        return selected

    def _on_success(self, result: dict):
        if self._on_merged:
            self._on_merged(result)
        self.destroy()

    def _on_error(self, err: str):
        self._btn_merge.configure(state="normal", text="Merge Customers")
        self._lbl_error.configure(text=f"Error: {err[:120]}")

    @staticmethod
    def _default_choice(value_a, value_b) -> str:
        if _is_empty(value_a) and not _is_empty(value_b):
            return "B"
        return "A"

    def _customer_ref(self, customer: dict) -> str:
        return str(customer.get("customer_code") or customer.get("customer_id") or "Unknown")

    def _customer_heading(self, customer: dict) -> str:
        name = f"{customer.get('first_name') or ''} {customer.get('surname') or ''}".strip()
        business = (customer.get("business") or "").strip()
        if name and business:
            return f"{name}\n{business}"
        return name or business or "(No name)"

    def _format_value(self, field_key: str, value) -> str:
        if value is None or value == "":
            return "-"
        if field_key == "terms_days":
            return f"{int(value)} days" if value else "-"
        if field_key == "credit_limit":
            try:
                return f"${float(value):,.2f}"
            except Exception:
                return str(value)
        if field_key in {"ship_same_as_invoice", "stop_credit", "is_local", "newsletter_opt_in"}:
            return "Yes" if bool(value) else "No"
        return str(value)


def _is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False
