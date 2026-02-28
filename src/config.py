from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Optional

CONFIG_PATH = Path("config.json")
EXAMPLE_PATH = Path("config.example.json")


@dataclass
class SupplierConfig:
    name: str
    suffix: str
    suffix_position: Literal["append", "prepend"]
    character_substitutions: dict
    pdf_format: Literal["table", "text", "marker"]
    validation_marker: str
    # --- table mode ---
    sku_column_hint: str = ""
    qty_column_hint: str = ""
    desc_column_hint: str = ""
    # --- text mode (regex) ---
    sku_pattern: str = ""
    qty_pattern: str = ""
    # --- marker mode ---
    item_start_marker: str = ""      # text marking the start of the items section
    item_end_marker: str = ""        # text marking the end of the items section
    trailing_numeric_count: int = 4  # numeric fields AFTER qty (e.g. RRP, UnitPrice, Discount%, Amount)
    sku_min_digits: int = 5          # minimum digit-only token length to be treated as a SKU


@dataclass
class NetoConfig:
    store_url: str
    api_key: str
    username: str


@dataclass
class EbayConfig:
    client_id: str
    client_secret: str
    ru_name: str
    refresh_token: str
    access_token: str
    access_token_expires_at: float
    environment: Literal["production", "sandbox"]


@dataclass
class AppConfig:
    order_lookback_days: int
    on_po_filter_phrase: str
    output_dir: str


class ConfigManager:
    def __init__(self):
        self._raw: dict = {}
        self.neto: NetoConfig = None
        self.ebay: EbayConfig = None
        self.suppliers: list[SupplierConfig] = []
        self.app: AppConfig = None

    def load(self) -> None:
        if not CONFIG_PATH.exists():
            raise FileNotFoundError(str(CONFIG_PATH))
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            self._raw = json.load(f)
        self._parse()

    def _parse(self) -> None:
        n = self._raw["neto"]
        self.neto = NetoConfig(
            store_url=n["store_url"].rstrip("/"),
            api_key=n["api_key"],
            username=n.get("username", ""),
        )

        e = self._raw["ebay"]
        self.ebay = EbayConfig(
            client_id=e.get("client_id", ""),
            client_secret=e.get("client_secret", ""),
            ru_name=e.get("ru_name", ""),
            refresh_token=e.get("refresh_token", ""),
            access_token=e.get("access_token", ""),
            access_token_expires_at=float(e.get("access_token_expires_at", 0)),
            environment=e.get("environment", "production"),
        )

        self.suppliers = [
            SupplierConfig(
                name=s["name"],
                suffix=s.get("suffix", ""),
                suffix_position=s.get("suffix_position", "append"),
                character_substitutions=s.get("character_substitutions", {}),
                pdf_format=s.get("pdf_format", "table"),
                validation_marker=s.get("validation_marker", ""),
                sku_column_hint=s.get("sku_column_hint", ""),
                qty_column_hint=s.get("qty_column_hint", ""),
                desc_column_hint=s.get("desc_column_hint", ""),
                sku_pattern=s.get("sku_pattern", ""),
                qty_pattern=s.get("qty_pattern", ""),
                item_start_marker=s.get("item_start_marker", ""),
                item_end_marker=s.get("item_end_marker", ""),
                trailing_numeric_count=int(s.get("trailing_numeric_count", 4)),
                sku_min_digits=int(s.get("sku_min_digits", 5)),
            )
            for s in self._raw.get("suppliers", [])
        ]

        a = self._raw.get("app", {})
        self.app = AppConfig(
            order_lookback_days=a.get("order_lookback_days", 30),
            on_po_filter_phrase=a.get("on_po_filter_phrase", "on po"),
            output_dir=a.get("output_dir", "output"),
        )

    def save(self) -> None:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(self._raw, f, indent=2, ensure_ascii=False)

    def save_ebay_tokens(
        self,
        access_token: str,
        expires_at: float,
        refresh_token: str = None,
    ) -> None:
        self._raw["ebay"]["access_token"] = access_token
        self._raw["ebay"]["access_token_expires_at"] = expires_at
        if refresh_token is not None:
            self._raw["ebay"]["refresh_token"] = refresh_token
        self.save()
        # Update in-memory state
        self.ebay.access_token = access_token
        self.ebay.access_token_expires_at = expires_at
        if refresh_token is not None:
            self.ebay.refresh_token = refresh_token

    def get_supplier_by_name(self, name: str) -> SupplierConfig | None:
        for s in self.suppliers:
            if s.name == name:
                return s
        return None

    def supplier_names(self) -> list[str]:
        return [s.name for s in self.suppliers]


# Module-level singleton
config = ConfigManager()
