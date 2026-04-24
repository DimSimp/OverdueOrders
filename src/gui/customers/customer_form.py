from __future__ import annotations

import threading
from typing import Callable, Optional

import customtkinter as ctk

from src.customers.discount_profiles import DISCOUNT_PROFILE_OPTIONS


_TERMS_OPTIONS = ["-", "7 days", "14 days", "30 days", "60 days"]
_TERMS_MAP = {"-": None, "7 days": 7, "14 days": 14, "30 days": 30, "60 days": 60}
_TERMS_REVERSE = {v: k for k, v in _TERMS_MAP.items()}


class CustomerForm(ctk.CTkToplevel):
    """New / Edit customer modal dialog."""

    def __init__(
        self,
        master,
        customer: Optional[dict] = None,
        on_saved: Optional[Callable[[dict], None]] = None,
    ):
        super().__init__(master)
        self._customer = customer
        self._on_saved = on_saved
        self._is_edit = customer is not None

        title = (
            f"Edit Customer - {customer.get('first_name', '')} {customer.get('surname', '')}".strip()
            if self._is_edit else "New Customer"
        )
        self.title(title)
        self.geometry("980x820")
        self.minsize(860, 700)
        self.resizable(True, True)

        self.transient(master)
        self.grab_set()

        self._build_ui()
        if self._is_edit:
            self._populate(customer)

        self.after(50, self._focus_first)

    # Build

    def _build_ui(self):
        self._scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._scroll.pack(fill="both", expand=True, padx=0, pady=0)
        body = self._scroll

        top = ctk.CTkFrame(body, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=(16, 10))
        top.grid_columnconfigure(0, weight=1, uniform="customer_form_top")
        top.grid_columnconfigure(1, weight=1, uniform="customer_form_top")

        left_col = ctk.CTkFrame(top, fg_color="transparent")
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        right_col = ctk.CTkFrame(top, fg_color="transparent")
        right_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        general = self._section_card(left_col, "General Details")
        general.grid_columnconfigure(0, weight=1)
        general.grid_columnconfigure(1, weight=1)

        self._e_first_name = self._field_block(general, "First Name *", row=0, column=0)
        self._e_surname = self._field_block(general, "Surname", row=0, column=1)
        self._e_business = self._field_block(
            general, "Business / School", row=1, column=0, columnspan=2
        )
        self._e_mobile = self._field_block(general, "Mobile", row=2, column=0)
        self._e_phone_1 = self._field_block(general, "Phone", row=2, column=1)
        self._e_email = self._field_block(general, "Email", row=3, column=0, columnspan=2)

        invoice = self._section_card(right_col, "Invoice Address")
        self._build_address_section(
            invoice,
            prefix="invoice",
            defaults={"country": "Australia"},
        )

        shipping = self._section_card(
            right_col,
            "Shipping Address",
            action_text="Copy Invoice -> Shipping",
            action_command=self._copy_invoice_to_shipping,
        )
        self._build_address_section(
            shipping,
            prefix="shipping",
            defaults={"country": "Australia"},
        )

        account = self._section_card(body, "Account Details", outer_padx=16)
        account.grid_columnconfigure(0, weight=1)
        account.grid_columnconfigure(1, weight=1)

        left_account = ctk.CTkFrame(account, fg_color="transparent")
        left_account.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left_account.grid_columnconfigure(0, weight=1)

        right_account = ctk.CTkFrame(account, fg_color="transparent")
        right_account.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        right_account.grid_columnconfigure(0, weight=1)

        self._opt_terms = self._option_block(
            left_account,
            "Payment Terms",
            _TERMS_OPTIONS,
            row=0,
            column=0,
        )
        self._opt_terms.set("-")
        self._opt_discount_profile = self._option_block(
            left_account,
            "Profile Discount",
            DISCOUNT_PROFILE_OPTIONS,
            row=1,
            column=0,
        )
        self._opt_discount_profile.set("-")
        self._e_credit_limit = self._field_block(left_account, "Credit Limit ($)", row=2, column=0)
        self._e_abn = self._field_block(left_account, "ABN", row=3, column=0)
        self._e_tax_exempt = self._field_block(
            left_account, "Tax Exemption #", row=4, column=0
        )

        toggles = ctk.CTkFrame(right_account, fg_color="transparent")
        toggles.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._sw_stop_credit = self._toggle(toggles, "Stop Credit")
        self._sw_is_local = self._toggle(toggles, "Is Local")
        self._sw_newsletter = self._toggle(toggles, "Newsletter Opt-in")
        self._sw_active = self._toggle(toggles, "Active", default=True)

        self._txt_private = self._textbox_block(
            right_account,
            "Private Comment",
            row=1,
            column=0,
            height=92,
        )
        self._txt_statement = self._textbox_block(
            right_account,
            "Statement Comment",
            row=2,
            column=0,
            height=92,
        )

        footer = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), height=52)
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

        self._btn_save = ctk.CTkButton(
            footer,
            text="Save Customer",
            width=130,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._on_save,
        )
        self._btn_save.pack(side="right", padx=(0, 8), pady=10)

    # UI helpers

    def _section_card(
        self,
        parent,
        title: str,
        action_text: str | None = None,
        action_command=None,
        outer_padx=0,
        outer_pady=(0, 16),
    ):
        card = ctk.CTkFrame(
            parent,
            fg_color=("gray92", "gray17"),
            corner_radius=10,
            border_width=1,
            border_color=("gray78", "gray28"),
        )
        card.pack(fill="x", padx=outer_padx, pady=outer_pady)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(12, 8))

        ctk.CTkLabel(
            header,
            text=title,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).pack(side="left")

        if action_text and action_command:
            ctk.CTkButton(
                header,
                text=action_text,
                width=170,
                height=28,
                font=ctk.CTkFont(size=11),
                fg_color="transparent",
                border_width=1,
                border_color=("gray60", "gray45"),
                text_color=("gray30", "gray70"),
                hover_color=("gray85", "gray25"),
                command=action_command,
            ).pack(side="right")

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        return content

    def _build_address_section(self, parent, prefix: str, defaults: dict | None = None):
        defaults = defaults or {}
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)

        setattr(
            self,
            f"_e_{prefix}_address_1",
            self._field_block(parent, "Address Line 1", row=0, column=0, columnspan=2),
        )
        setattr(
            self,
            f"_e_{prefix}_address_2",
            self._field_block(parent, "Address Line 2", row=1, column=0, columnspan=2),
        )
        setattr(self, f"_e_{prefix}_city", self._field_block(parent, "City", row=2, column=0))
        setattr(self, f"_e_{prefix}_state", self._field_block(parent, "State", row=2, column=1))
        setattr(
            self,
            f"_e_{prefix}_postcode",
            self._field_block(parent, "Postcode", row=3, column=0),
        )
        setattr(
            self,
            f"_e_{prefix}_country",
            self._field_block(
                parent,
                "Country",
                row=3,
                column=1,
                default=defaults.get("country", ""),
            ),
        )

    def _field_block(
        self,
        parent,
        label: str,
        *,
        row: int,
        column: int,
        columnspan: int = 1,
        default: str = "",
    ) -> ctk.CTkEntry:
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=4, pady=4)

        ctk.CTkLabel(
            wrap,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray70"),
            anchor="w",
        ).pack(fill="x", pady=(0, 4))

        entry = ctk.CTkEntry(wrap, font=ctk.CTkFont(size=12))
        entry.pack(fill="x")
        if default:
            entry.insert(0, default)
        return entry

    def _option_block(
        self,
        parent,
        label: str,
        values: list[str],
        *,
        row: int,
        column: int,
    ) -> ctk.CTkOptionMenu:
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.grid(row=row, column=column, sticky="ew", padx=4, pady=4)

        ctk.CTkLabel(
            wrap,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray70"),
            anchor="w",
        ).pack(fill="x", pady=(0, 4))

        opt = ctk.CTkOptionMenu(
            wrap,
            values=values,
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        opt.pack(fill="x")
        return opt

    def _textbox_block(
        self,
        parent,
        label: str,
        *,
        row: int,
        column: int,
        height: int = 72,
    ) -> ctk.CTkTextbox:
        wrap = ctk.CTkFrame(parent, fg_color="transparent")
        wrap.grid(row=row, column=column, sticky="ew", padx=4, pady=4)

        ctk.CTkLabel(
            wrap,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray70"),
            anchor="w",
        ).pack(fill="x", pady=(0, 4))

        tb = ctk.CTkTextbox(wrap, height=height, font=ctk.CTkFont(size=12))
        tb.pack(fill="x")
        return tb

    def _toggle(self, parent, label: str, default: bool = False) -> ctk.CTkSwitch:
        sw = ctk.CTkSwitch(parent, text=label, font=ctk.CTkFont(size=11))
        sw.pack(anchor="w", pady=4)
        if default:
            sw.select()
        return sw

    # Populate

    def _populate(self, c: dict):
        self._set(self._e_first_name, c.get("first_name"))
        self._set(self._e_surname, c.get("surname"))
        self._set(self._e_business, c.get("business"))
        self._set(self._e_mobile, c.get("mobile"))
        self._set(self._e_phone_1, c.get("phone_1"))
        self._set(self._e_email, c.get("email"))

        self._set(self._e_invoice_address_1, c.get("address_1"))
        self._set(self._e_invoice_address_2, c.get("address_2"))
        self._set(self._e_invoice_city, c.get("city"))
        self._set(self._e_invoice_state, c.get("state"))
        self._set(self._e_invoice_postcode, c.get("postcode"))
        self._set(self._e_invoice_country, c.get("country") or "Australia")

        self._set(self._e_shipping_address_1, c.get("ship_address_1"))
        self._set(self._e_shipping_address_2, c.get("ship_address_2"))
        self._set(self._e_shipping_city, c.get("ship_city"))
        self._set(self._e_shipping_state, c.get("ship_state"))
        self._set(self._e_shipping_postcode, c.get("ship_postcode"))
        self._set(self._e_shipping_country, c.get("ship_country") or "Australia")

        if c.get("ship_same_as_invoice") and not any(
            [
                c.get("ship_address_1"),
                c.get("ship_address_2"),
                c.get("ship_city"),
                c.get("ship_state"),
                c.get("ship_postcode"),
                c.get("ship_country"),
            ]
        ):
            self._copy_invoice_to_shipping()

        self._set(self._e_credit_limit, c.get("credit_limit"))
        self._opt_discount_profile.set(c.get("discount_profile") or "-")
        self._set(self._e_abn, c.get("abn"))
        self._set(self._e_tax_exempt, c.get("tax_exemption_number"))

        self._opt_terms.set(_TERMS_REVERSE.get(c.get("terms_days"), "-"))

        if c.get("stop_credit"):
            self._sw_stop_credit.select()
        if c.get("is_local"):
            self._sw_is_local.select()
        if c.get("newsletter_opt_in"):
            self._sw_newsletter.select()
        if c.get("active", True):
            self._sw_active.select()
        else:
            self._sw_active.deselect()

        if c.get("private_comment"):
            self._txt_private.insert("1.0", c["private_comment"])
        if c.get("statement_comment"):
            self._txt_statement.insert("1.0", c["statement_comment"])

    @staticmethod
    def _set(entry: ctk.CTkEntry, value):
        if value is not None:
            entry.delete(0, "end")
            entry.insert(0, str(value))

    # Address helpers

    def _copy_invoice_to_shipping(self):
        pairs = [
            (self._e_invoice_address_1, self._e_shipping_address_1),
            (self._e_invoice_address_2, self._e_shipping_address_2),
            (self._e_invoice_city, self._e_shipping_city),
            (self._e_invoice_state, self._e_shipping_state),
            (self._e_invoice_postcode, self._e_shipping_postcode),
            (self._e_invoice_country, self._e_shipping_country),
        ]
        for source, target in pairs:
            target.delete(0, "end")
            target.insert(0, source.get().strip())

    def _address_data(self, prefix: str) -> dict:
        return {
            "address_1": getattr(self, f"_e_{prefix}_address_1").get().strip() or None,
            "address_2": getattr(self, f"_e_{prefix}_address_2").get().strip() or None,
            "city": getattr(self, f"_e_{prefix}_city").get().strip() or None,
            "state": getattr(self, f"_e_{prefix}_state").get().strip() or None,
            "postcode": getattr(self, f"_e_{prefix}_postcode").get().strip() or None,
            "country": getattr(self, f"_e_{prefix}_country").get().strip() or "Australia",
        }

    @staticmethod
    def _addresses_match(invoice: dict, shipping: dict) -> bool:
        keys = ("address_1", "address_2", "city", "state", "postcode", "country")
        return all((invoice.get(k) or "") == (shipping.get(k) or "") for k in keys)

    # Save

    def _on_save(self):
        self._lbl_error.configure(text="")

        first = self._e_first_name.get().strip()
        if not first:
            self._lbl_error.configure(text="First Name is required.")
            self._e_first_name.focus_set()
            return

        mobile = self._e_mobile.get().strip()
        phone_1 = self._e_phone_1.get().strip()
        if not mobile and not phone_1:
            self._lbl_error.configure(text="Either Mobile or Phone is required.")
            self._e_mobile.focus_set()
            return

        invoice = self._address_data("invoice")
        shipping = self._address_data("shipping")

        data: dict = {
            "first_name": first or None,
            "surname": self._e_surname.get().strip() or None,
            "business": self._e_business.get().strip() or None,
            "mobile": mobile or None,
            "phone_1": phone_1 or None,
            "email": self._e_email.get().strip() or None,
            "address_1": invoice["address_1"],
            "address_2": invoice["address_2"],
            "city": invoice["city"],
            "state": invoice["state"],
            "postcode": invoice["postcode"],
            "country": invoice["country"],
            "ship_address_1": shipping["address_1"],
            "ship_address_2": shipping["address_2"],
            "ship_city": shipping["city"],
            "ship_state": shipping["state"],
            "ship_postcode": shipping["postcode"],
            "ship_country": shipping["country"],
            "ship_same_as_invoice": self._addresses_match(invoice, shipping),
            "discount_profile": self._opt_discount_profile.get()
            if self._opt_discount_profile.get() != "-" else None,
            "credit_limit": self._parse_decimal(self._e_credit_limit.get()),
            "abn": self._e_abn.get().strip() or None,
            "tax_exemption_number": self._e_tax_exempt.get().strip() or None,
            "terms_days": _TERMS_MAP.get(self._opt_terms.get()),
            "stop_credit": self._sw_stop_credit.get() == 1,
            "is_local": self._sw_is_local.get() == 1,
            "newsletter_opt_in": self._sw_newsletter.get() == 1,
            "active": self._sw_active.get() == 1,
            "private_comment": self._txt_private.get("1.0", "end").strip() or None,
            "statement_comment": self._txt_statement.get("1.0", "end").strip() or None,
        }

        self._btn_save.configure(state="disabled", text="Saving...")

        def _thread():
            from src.customers.customer_client import create_customer, update_customer

            try:
                if self._is_edit:
                    saved = update_customer(self._customer["id"], data)
                else:
                    saved = create_customer(data)
                self.after(0, lambda: self._on_success(saved))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._on_error(err))

        threading.Thread(target=_thread, daemon=True).start()

    def _on_success(self, saved: dict):
        if self._on_saved:
            self._on_saved(saved)
        self.destroy()

    def _on_error(self, err: str):
        self._btn_save.configure(state="normal", text="Save Customer")
        self._lbl_error.configure(text=f"Error: {err[:80]}")

    @staticmethod
    def _parse_decimal(val: str) -> Optional[float]:
        val = val.strip().replace(",", "").replace("$", "")
        if not val:
            return None
        try:
            return float(val)
        except ValueError:
            return None

    def _focus_first(self):
        self._e_first_name.focus_set()
