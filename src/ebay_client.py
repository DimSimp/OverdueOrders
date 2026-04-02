from __future__ import annotations

import base64
import logging
import math
import time
import webbrowser
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timezone
from urllib.parse import urlencode, parse_qs, urlparse

import requests

log = logging.getLogger("ebay_client")

from src.config import EbayConfig

EBAY_PROD_TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token"
EBAY_PROD_AUTH_URL = "https://auth.ebay.com/oauth2/authorize"
EBAY_PROD_API_BASE = "https://api.ebay.com"

EBAY_SANDBOX_TOKEN_URL = "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
EBAY_SANDBOX_AUTH_URL = "https://auth.sandbox.ebay.com/oauth2/authorize"
EBAY_SANDBOX_API_BASE = "https://api.sandbox.ebay.com"

EBAY_AU_MARKETPLACE_ID = "EBAY_AU"
EBAY_SCOPE = "https://api.ebay.com/oauth/api_scope/sell.fulfillment"

# Trading API (SOAP) — for reading PrivateNotes via GetSellerTransactions
EBAY_PROD_TRADING_URL = "https://api.ebay.com/ws/api.dll"
EBAY_SANDBOX_TRADING_URL = "https://api.sandbox.ebay.com/ws/api.dll"
EBAY_AU_SITE_ID = "15"
EBAY_TRADING_VERSION = "967"
_TRADING_NS = "urn:ebay:apis:eBLBaseComponents"


@dataclass
class EbayLineItem:
    line_item_id: str
    sku: str
    title: str
    quantity: int
    legacy_item_id: str = ""        # Trading API ItemID (for PrivateNotes lookup)
    legacy_transaction_id: str = "" # Trading API TransactionID
    notes: str = ""                 # PrivateNotes for this specific item
    image_url: str = ""             # Product image URL
    unit_price: float = 0.0         # Price per unit
    variation_specifics: str = ""   # Raw variation string when no SKU is set on the listing
                                    # e.g. "Size: Medium Light ·String: 3G - .022"


@dataclass
class EbayOrder:
    order_id: str
    buyer_name: str
    buyer_notes: str
    creation_date: datetime | None
    order_status: str       # orderFulfillmentStatus
    payment_status: str     # orderPaymentStatus
    line_items: list[EbayLineItem] = field(default_factory=list)
    # Shipping address
    ship_name: str = ""
    ship_street1: str = ""
    ship_street2: str = ""
    ship_city: str = ""
    ship_state: str = ""
    ship_postcode: str = ""
    ship_country: str = ""
    ship_phone: str = ""
    # Pricing & shipping
    order_total: float = 0.0
    shipping_cost: float = 0.0
    shipping_method: str = ""
    shipping_type: str = ""  # "Express", "Regular", "Local Pickup", or ""
    # Tracking (populated for fulfilled orders)
    tracking_number: str = ""
    tracking_carrier: str = ""


class EbayAuthError(Exception):
    pass


class EbayAPIError(Exception):
    pass


class EbayClient:
    def __init__(self, config: EbayConfig, token_save_callback, variation_sku_manager=None):
        """
        token_save_callback(access_token, expires_at, refresh_token=None)
        called after any token exchange or refresh to persist tokens.

        variation_sku_manager: optional EbayVariationSkuManager instance; when
        provided, variation specifics strings are resolved to real SKUs via the
        CSV cache and/or Trading API GetItem.
        """
        self._config = config
        self._save_tokens = token_save_callback
        self._variation_sku_manager = variation_sku_manager
        self._session = requests.Session()
        self.notes_warning: str = ""  # set by _enrich_with_private_notes on failure

    # ----- Auth helpers -----

    @property
    def _token_url(self) -> str:
        return EBAY_SANDBOX_TOKEN_URL if self._config.environment == "sandbox" else EBAY_PROD_TOKEN_URL

    @property
    def _auth_url(self) -> str:
        return EBAY_SANDBOX_AUTH_URL if self._config.environment == "sandbox" else EBAY_PROD_AUTH_URL

    @property
    def _api_base(self) -> str:
        return EBAY_SANDBOX_API_BASE if self._config.environment == "sandbox" else EBAY_PROD_API_BASE

    def _credentials_header(self) -> str:
        encoded = base64.b64encode(
            f"{self._config.client_id}:{self._config.client_secret}".encode()
        ).decode()
        return f"Basic {encoded}"

    def is_authenticated(self) -> bool:
        return bool(self._config.refresh_token)

    def get_auth_url(self) -> str:
        """Build the eBay OAuth consent URL for the user to open in a browser."""
        params = {
            "client_id": self._config.client_id,
            "redirect_uri": self._config.ru_name,
            "response_type": "code",
            "scope": EBAY_SCOPE,
        }
        return f"{self._auth_url}?{urlencode(params)}"

    def open_auth_in_browser(self) -> None:
        webbrowser.open(self.get_auth_url())

    def exchange_code(self, redirect_url_or_code: str) -> None:
        """
        Accept either the full redirect URL (pasted from browser address bar)
        or a raw authorization code, then exchange for tokens.
        """
        code = _extract_code(redirect_url_or_code)
        if not code:
            raise EbayAuthError(
                "Could not find an authorization code in the URL you provided.\n"
                "Make sure you pasted the full URL from the browser address bar after approving access."
            )

        resp = self._session.post(
            self._token_url,
            headers={
                "Authorization": self._credentials_header(),
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data={
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": self._config.ru_name,
            },
            timeout=30,
        )
        self._raise_for_ebay_error(resp)
        data = resp.json()
        expires_at = time.time() + data["expires_in"] - 60
        self._save_tokens(data["access_token"], expires_at, data.get("refresh_token", ""))

    def _ensure_valid_token(self) -> str:
        """Return a valid access token, auto-refreshing if expired."""
        if not self._config.refresh_token:
            raise EbayAuthError(
                "eBay is not authenticated.\n"
                "Please use the 'Authenticate eBay' button in the Orders tab to set up access."
            )
        if time.time() < self._config.access_token_expires_at and self._config.access_token:
            return self._config.access_token
        return self._refresh_access_token()

    def _refresh_access_token(self) -> str:
        resp = self._session.post(
            self._token_url,
            headers={
                "Authorization": self._credentials_header(),
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data={
                "grant_type": "refresh_token",
                "refresh_token": self._config.refresh_token,
                "scope": EBAY_SCOPE,
            },
            timeout=30,
        )
        self._raise_for_ebay_error(resp)
        data = resp.json()
        expires_at = time.time() + data["expires_in"] - 60
        self._save_tokens(data["access_token"], expires_at)
        return data["access_token"]

    @staticmethod
    def _raise_for_ebay_error(resp: requests.Response) -> None:
        try:
            resp.raise_for_status()
        except requests.HTTPError as e:
            try:
                detail = resp.json()
                msg = detail.get("error_description") or detail.get("message") or str(detail)
            except Exception:
                msg = resp.text
            raise EbayAPIError(f"eBay API error ({resp.status_code}): {msg}") from e

    # ----- Orders -----

    def get_overdue_orders(
        self,
        date_from: datetime,
        date_to: datetime,
        progress_callback=None,
    ) -> list[EbayOrder]:
        """
        Fetch paid orders with unfulfilled status within the date range.
        Handles pagination automatically.
        """
        token = self._ensure_valid_token()

        from_str = _to_ebay_datetime(date_from)
        to_str = _to_ebay_datetime(date_to)
        date_filter = f"creationdate:[{from_str}..{to_str}]"
        status_filter = "orderfulfillmentstatus:{NOT_STARTED|IN_PROGRESS}"
        combined_filter = f"{date_filter},{status_filter}"
        log.debug("eBay fetch: date_from=%s  date_to=%s  filter=%s", date_from, date_to, combined_filter)

        all_raw: list[dict] = []
        offset = 0
        limit = 200
        total = None

        while True:
            resp = self._session.get(
                f"{self._api_base}/sell/fulfillment/v1/order",
                headers={
                    "Authorization": f"Bearer {token}",
                    "X-EBAY-C-MARKETPLACE-ID": EBAY_AU_MARKETPLACE_ID,
                    "Content-Type": "application/json",
                },
                params={
                    "filter": combined_filter,
                    "limit": limit,
                    "offset": offset,
                },
                timeout=30,
            )
            self._raise_for_ebay_error(resp)
            data = resp.json()

            if total is None:
                total = data.get("total", 0)
                log.debug("eBay fetch: API reports total=%d matching orders", total)

            batch = data.get("orders", [])
            all_raw.extend(batch)
            log.debug("eBay fetch: page offset=%d  got %d orders  running total=%d", offset, len(batch), len(all_raw))

            if progress_callback:
                progress_callback(len(all_raw), total or len(all_raw))

            if not batch or len(all_raw) >= total:
                break
            offset += limit

        # Filter to PAID orders (eBay fulfillment filter doesn't support payment status)
        paid = [o for o in all_raw if o.get("orderPaymentStatus") == "PAID"]
        unpaid = [o for o in all_raw if o.get("orderPaymentStatus") != "PAID"]
        if unpaid:
            log.debug(
                "eBay fetch: %d orders excluded (not PAID) — statuses: %s",
                len(unpaid),
                [o.get("orderPaymentStatus") for o in unpaid],
            )
        log.debug("eBay fetch: %d raw orders → %d after PAID filter", len(all_raw), len(paid))

        orders = [self._parse_order(o) for o in paid]
        for o in orders:
            skus = [li.sku for li in o.line_items]
            log.debug(
                "  order %-24s | %-30s | created %-20s | status=%-12s | payment=%-6s | skus=%s | checkout_notes=%r",
                o.order_id, o.buyer_name,
                str(o.creation_date)[:19] if o.creation_date else "—",
                o.order_status, o.payment_status,
                skus, o.buyer_notes,
            )

        # Enrich PrivateNotes from Trading API (also populates variation_specifics
        # for listings where Fulfillment API returns no SKU/lineItemProperties)
        if self._config.dev_id:
            self._enrich_with_private_notes(orders, date_from, date_to)
        else:
            log.warning(
                "eBay Trading API credentials not configured (dev_id missing) — "
                "PrivateNotes will not be fetched. Add dev_id to config.json to enable."
            )

        # Resolve variation specifics strings to real SKUs (run after Trading API
        # enrichment so variation_specifics is populated even when lineItemProperties
        # was absent from the Fulfillment API response)
        self._enrich_variation_skus(orders)

        log.debug("eBay fetch complete: returning %d orders", len(orders))
        for o in orders:
            if o.buyer_notes:
                log.debug("  order %s  final_notes=%r", o.order_id, o.buyer_notes)

        return orders

    def get_orders_by_ids(self, order_ids: list[str]) -> list[EbayOrder]:
        """Fetch specific orders by ID. Returns only paid, unfulfilled orders.

        eBay's Fulfillment API accepts at most 50 order IDs per request, so
        larger lists are automatically split into batches.
        """
        if not order_ids:
            return []
        token = self._ensure_valid_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "X-EBAY-C-MARKETPLACE-ID": EBAY_AU_MARKETPLACE_ID,
            "Content-Type": "application/json",
        }

        _BATCH_SIZE = 50
        raw_orders: list = []
        for i in range(0, len(order_ids), _BATCH_SIZE):
            batch = order_ids[i: i + _BATCH_SIZE]
            resp = self._session.get(
                f"{self._api_base}/sell/fulfillment/v1/order",
                headers=headers,
                params={"orderIds": ",".join(batch)},
                timeout=30,
            )
            self._raise_for_ebay_error(resp)
            raw_orders.extend(resp.json().get("orders", []))

        filtered = [
            o for o in raw_orders
            if o.get("orderPaymentStatus") == "PAID"
            and o.get("orderFulfillmentStatus") in ("NOT_STARTED", "IN_PROGRESS")
        ]
        orders = [self._parse_order(o) for o in filtered]
        if orders and self._config.dev_id:
            # Use a wide window (60 days) since we don't have an explicit date_from
            from datetime import timedelta
            date_from = datetime.now() - timedelta(days=60)
            date_to = datetime.now()
            self._enrich_with_private_notes(orders, date_from, date_to)
        self._enrich_variation_skus(orders)
        return orders

    def search_orders(
        self,
        date_from: datetime,
        date_to: datetime,
        fulfillment_statuses: list[str] | None = None,
        enrich_notes: bool = False,
        progress_callback=None,
    ) -> list[EbayOrder]:
        """
        Fetch eBay orders by creation date range and optional fulfillment status filter.
        No payment status filter — all payment states are returned.
        fulfillment_statuses: e.g. ["NOT_STARTED","IN_PROGRESS"] or ["FULFILLED"].
          Pass None to omit the status filter entirely.
        enrich_notes: if True, fetches PrivateNotes via Trading API (slow). False by default
          for search — notes still appear when opening the full order detail view.
        """
        token = self._ensure_valid_token()
        from_str = _to_ebay_datetime(date_from)
        to_str = _to_ebay_datetime(date_to)
        date_filter = f"creationdate:[{from_str}..{to_str}]"
        if fulfillment_statuses:
            status_str = "|".join(fulfillment_statuses)
            combined_filter = f"{date_filter},orderfulfillmentstatus:{{{status_str}}}"
        else:
            combined_filter = date_filter

        all_raw: list[dict] = []
        offset = 0
        limit = 200
        total = None

        while True:
            resp = self._session.get(
                f"{self._api_base}/sell/fulfillment/v1/order",
                headers={
                    "Authorization": f"Bearer {token}",
                    "X-EBAY-C-MARKETPLACE-ID": EBAY_AU_MARKETPLACE_ID,
                    "Content-Type": "application/json",
                },
                params={
                    "filter": combined_filter,
                    "limit": limit,
                    "offset": offset,
                },
                timeout=30,
            )
            self._raise_for_ebay_error(resp)
            data = resp.json()

            if total is None:
                total = data.get("total", 0)

            batch = data.get("orders", [])
            all_raw.extend(batch)

            if progress_callback:
                progress_callback(len(all_raw), total or len(all_raw))

            if not batch or len(all_raw) >= (total or 0):
                break
            offset += limit

        orders = [self._parse_order(o) for o in all_raw]
        if enrich_notes and self._config.dev_id and orders:
            self._enrich_with_private_notes(orders, date_from, date_to)
        self._enrich_variation_skus(orders)
        log.debug("eBay search_orders: returning %d orders", len(orders))
        return orders

    def get_order_by_exact_id(self, order_id: str) -> list[EbayOrder]:
        """Fetch a single eBay order by exact order ID. Returns empty list if not found."""
        token = self._ensure_valid_token()
        resp = self._session.get(
            f"{self._api_base}/sell/fulfillment/v1/order/{order_id}",
            headers={
                "Authorization": f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": EBAY_AU_MARKETPLACE_ID,
                "Content-Type": "application/json",
            },
            timeout=30,
        )
        if resp.status_code == 404:
            return []
        self._raise_for_ebay_error(resp)
        data = resp.json()
        orders = [self._parse_order(data)]
        self._enrich_variation_skus(orders)
        return orders

    def get_order_status(self, order_id: str) -> str:
        """Lightweight call to check current fulfillment status of an order."""
        token = self._ensure_valid_token()
        resp = self._session.get(
            f"{self._api_base}/sell/fulfillment/v1/order/{order_id}",
            headers={
                "Authorization": f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": EBAY_AU_MARKETPLACE_ID,
                "Content-Type": "application/json",
            },
            timeout=10,
        )
        self._raise_for_ebay_error(resp)
        data = resp.json()
        return data.get("orderFulfillmentStatus", "")

    def create_shipping_fulfillment(
        self,
        order_id: str,
        line_items: list[EbayLineItem],
        tracking_number: str = "",
        carrier: str = "",
        dry_run: bool = True,
    ) -> dict:
        """Create a shipping fulfillment for an order. Returns API response."""
        fulfillment_lines = [
            {"lineItemId": li.line_item_id, "quantity": li.quantity}
            for li in line_items
        ]
        body = {"lineItems": fulfillment_lines}
        if tracking_number:
            body["trackingNumber"] = tracking_number
        if carrier:
            body["shippingCarrierCode"] = carrier

        if dry_run:
            print(f"[DRY RUN] eBay CreateShippingFulfillment on {order_id}: {body}")
            return {"fulfillmentId": "DRY_RUN", "DryRun": True}

        token = self._ensure_valid_token()
        resp = self._session.post(
            f"{self._api_base}/sell/fulfillment/v1/order/{order_id}/shipping_fulfillment",
            headers={
                "Authorization": f"Bearer {token}",
                "X-EBAY-C-MARKETPLACE-ID": EBAY_AU_MARKETPLACE_ID,
                "Content-Type": "application/json",
            },
            json=body,
            timeout=30,
        )
        self._raise_for_ebay_error(resp)
        # eBay returns 204 No Content on success — no body to parse
        if resp.status_code == 204 or not resp.content:
            return {}
        return resp.json()

    # ----- Trading API (PrivateNotes) -----

    @property
    def _trading_url(self) -> str:
        return EBAY_SANDBOX_TRADING_URL if self._config.environment == "sandbox" else EBAY_PROD_TRADING_URL

    def _call_trading_api(self, xml_template: str, call_name: str) -> ET.Element:
        """Make a Trading API call, injecting the best available token.

        Tries user_token first if configured.  If the response returns error 931
        (Auth token is invalid / expired), automatically retries with the OAuth
        access token.  Raises EbayAPIError on unrecoverable failure.

        *xml_template* must contain ``{TOKEN}`` as a placeholder for the token.
        """
        ns = {"e": _TRADING_NS}
        tokens_tried: list[tuple[str, str]] = []

        if self._config.user_token:
            tokens_tried.append(("user_token", self._config.user_token))
        try:
            tokens_tried.append(("oauth", self._ensure_valid_token()))
        except Exception:
            pass  # no OAuth token yet; user_token is the only hope

        last_detail = "no tokens available"
        for token_type, token in tokens_tried:
            xml_body = xml_template.replace("{TOKEN}", token)
            try:
                resp = self._session.post(
                    self._trading_url,
                    headers={
                        "X-EBAY-API-CALL-NAME": call_name,
                        "X-EBAY-API-APP-NAME": self._config.client_id,
                        "X-EBAY-API-DEV-NAME": self._config.dev_id,
                        "X-EBAY-API-CERT-NAME": self._config.client_secret,
                        "X-EBAY-API-SITEID": EBAY_AU_SITE_ID,
                        "X-EBAY-API-COMPATIBILITY-LEVEL": EBAY_TRADING_VERSION,
                        "Content-Type": "text/xml",
                    },
                    data=xml_body.encode("utf-8"),
                    timeout=30,
                )
                resp.raise_for_status()
            except Exception as exc:
                log.warning("Trading API HTTP error with %s: %s", token_type, exc)
                last_detail = str(exc)
                continue

            root = ET.fromstring(resp.text)
            ack = root.find("e:Ack", ns)
            if ack is not None and ack.text in ("Success", "Warning"):
                if token_type == "user_token":
                    log.debug("Trading API call %s succeeded with user_token", call_name)
                else:
                    log.debug("Trading API call %s succeeded with oauth token", call_name)
                return root

            # Extract error details
            errors = root.findall(".//e:Errors", ns)
            error_msgs = []
            error_codes = []
            for err in errors:
                code_el = err.find("e:ErrorCode", ns)
                msg_el = err.find("e:LongMessage", ns) or err.find("e:ShortMessage", ns)
                if code_el is not None and code_el.text:
                    error_codes.append(code_el.text)
                if msg_el is not None and msg_el.text:
                    code = code_el.text if code_el is not None else ""
                    error_msgs.append(f"[{code}] {msg_el.text}" if code else msg_el.text)
            last_detail = "; ".join(error_msgs) if error_msgs else "Ack=Failure (no detail)"

            if "931" in error_codes and token_type == "user_token":
                log.warning(
                    "Trading API: user_token expired or invalid (931), retrying with OAuth token"
                )
                continue  # retry with next token

            log.warning("Trading API %s failed with %s: %s", call_name, token_type, last_detail)
            break  # non-recoverable error

        raise EbayAPIError(f"Trading API {call_name} failed: {last_detail}")

    def _enrich_with_private_notes(
        self,
        orders: list[EbayOrder],
        date_from: datetime,
        date_to: datetime,
    ) -> None:
        """
        Fetch PrivateNotes from the Trading API (GetSellerTransactions) and populate
        li.notes on each EbayLineItem. Also populates legacy_transaction_id where missing.

        If any line item on an order has a PrivateNote, the combined notes are
        appended to buyer_notes so filter_on_po can detect them.
        """
        self.notes_warning = ""

        # GetMyeBaySelling uses DurationInDays, not a date range.
        # Calculate days since date_from; cap at 60 (API limit for sold list).
        now = datetime.now()
        duration_days = min(int((now - date_from).days) + 1, 60)

        # Collect the item IDs we actually need, so we can stop early once found.
        # GetMyeBaySelling returns newest-first, so recent orders are on early pages.
        target_item_ids: set[str] = {
            li.legacy_item_id
            for o in orders
            for li in o.line_items
            if li.legacy_item_id
        }

        # Build lookups keyed by (ItemID, TransactionID) so that multiple sales
        # of the same listing don't overwrite each other's notes.
        notes_by_txn: dict[tuple[str, str], str] = {}
        # ItemID → list of TransactionIDs seen (used to correct "0" txn IDs from
        # the Fulfillment API when there is only one transaction for that item).
        txn_ids_by_item: dict[str, list[str]] = {}
        # (ItemID, TransactionID) → (variation_specifics_str, variation_sku)
        # Populated from GetMyeBaySelling for variation listings.
        variation_by_txn: dict[tuple[str, str], tuple[str, str]] = {}
        ns = {"e": _TRADING_NS}
        page = 1
        while True:
            xml_template = self._build_sold_list_xml(duration_days, page)
            try:
                root = self._call_trading_api(xml_template, "GetMyeBaySelling")
            except EbayAPIError as exc:
                self.notes_warning = f"PrivateNotes unavailable: {exc}"
                break

            for txn in root.findall(".//e:Transaction", ns):
                oli_el = txn.find("e:OrderLineItemID", ns)
                if oli_el is None or not oli_el.text:
                    continue
                # OLI format: "ItemID-TransactionID"
                parts = oli_el.text.split("-", 1)
                item_id = parts[0]
                transaction_id = parts[1] if len(parts) > 1 else ""
                if item_id and transaction_id:
                    txn_ids_by_item.setdefault(item_id, [])
                    if transaction_id not in txn_ids_by_item[item_id]:
                        txn_ids_by_item[item_id].append(transaction_id)
                # PrivateNotes is at Transaction level (confirmed from legacy code)
                note = _xml_text(txn, "e:PrivateNotes", ns)
                if not note:
                    note = _xml_text(txn, "e:Item/e:PrivateNotes", ns)
                if item_id and transaction_id and note:
                    notes_by_txn[(item_id, transaction_id)] = note

                # Extract variation specifics for variation listings
                if item_id and transaction_id:
                    var_specs: dict[str, str] = {}
                    for nvl in txn.findall("e:Variation/e:VariationSpecifics/e:NameValueList", ns):
                        name = _xml_text(nvl, "e:Name", ns)
                        value = _xml_text(nvl, "e:Value", ns)
                        if name:
                            var_specs[name] = value
                    if var_specs:
                        var_sku_el = txn.find("e:Variation/e:SKU", ns)
                        var_sku = (var_sku_el.text or "").strip() if var_sku_el is not None else ""
                        specs_str = " \u00b7".join(f"{k}: {v}" for k, v in var_specs.items())
                        variation_by_txn[(item_id, transaction_id)] = (specs_str, var_sku)

            # Paginate through SoldList pages
            total_pages_el = root.find(".//e:SoldList/e:PaginationResult/e:TotalNumberOfPages", ns)
            total_pages = int(total_pages_el.text) if total_pages_el is not None and total_pages_el.text else 1

            # Stop early if we've found transaction IDs for all target orders —
            # all remaining pages contain older sold items we don't need
            found_items = set(txn_ids_by_item.keys())
            if target_item_ids and target_item_ids.issubset(found_items):
                log.debug(
                    "GetMyeBaySelling: all %d target items found on page %d/%d — stopping early",
                    len(target_item_ids), page, total_pages,
                )
                break

            if page >= total_pages:
                break
            page += 1

        log.debug(
            "Trading API enrichment complete: %d PrivateNotes found across %d items",
            len(notes_by_txn), len(txn_ids_by_item),
        )
        if notes_by_txn:
            for (item_id, txn_id), note in notes_by_txn.items():
                log.debug("  PrivateNote item_id=%s txn_id=%s  note=%r", item_id, txn_id, note)
        if txn_ids_by_item and not notes_by_txn:
            self.notes_warning = (
                "PrivateNotes not returned — eBay Trading token may be expired. "
                "Use 'Update Trading Token' to renew it."
            )
        if not notes_by_txn and not txn_ids_by_item:
            return

        # For orders where the Fulfillment API gave an empty txn_id and there are
        # multiple sold transactions for the same item, we can't guess which
        # transaction belongs to which order. Use GetOrders (Trading API) to resolve.
        unresolved_order_ids = list(dict.fromkeys(
            order.order_id
            for order in orders
            for li in order.line_items
            if li.legacy_item_id
            and not li.legacy_transaction_id
            and len(txn_ids_by_item.get(li.legacy_item_id, [])) != 1
        ))
        txn_by_order_item: dict[tuple[str, str], str] = {}
        if unresolved_order_ids:
            log.debug(
                "GetOrders fallback: resolving txn IDs for %d unresolved orders: %s",
                len(unresolved_order_ids), unresolved_order_ids,
            )
            txn_by_order_item = self._get_txn_ids_for_orders(unresolved_order_ids)

        # Apply notes per line item, keyed by (ItemID, TransactionID)
        for order in orders:
            item_notes = []
            for li in order.line_items:
                item_id = li.legacy_item_id
                txn_id = li.legacy_transaction_id

                log.debug(
                    "  enrichment lookup: order=%s  sku=%s  item_id=%s  txn_id=%s (from fulfillment API)",
                    order.order_id, li.sku, item_id, txn_id,
                )

                # Correct "0" or missing transaction IDs from the Fulfillment API.
                if item_id and (not txn_id or txn_id == "0"):
                    candidates = txn_ids_by_item.get(item_id, [])
                    log.debug(
                        "    txn_id is %r — candidates from sold list: %s",
                        txn_id, candidates,
                    )
                    if len(candidates) == 1:
                        txn_id = candidates[0]
                        li.legacy_transaction_id = txn_id
                        log.debug("    corrected txn_id → %s (single candidate)", txn_id)
                    else:
                        # Multiple (or zero) candidates — use GetOrders result
                        better = txn_by_order_item.get((order.order_id, item_id), "")
                        if better:
                            txn_id = better
                            li.legacy_transaction_id = txn_id
                            log.debug("    corrected txn_id → %s (GetOrders fallback)", txn_id)

                note = notes_by_txn.get((item_id, txn_id), "") if item_id and txn_id else ""
                log.debug(
                    "    lookup (%s, %s) → %s",
                    item_id, txn_id, repr(note) if note else "no match",
                )
                li.notes = note
                if note:
                    item_notes.append(note)
            if item_notes:
                combined = " | ".join(item_notes)
                if order.buyer_notes:
                    order.buyer_notes = order.buyer_notes + " | " + combined
                else:
                    order.buyer_notes = combined

        # Apply variation specifics from the Trading API sold list for items that
        # have no SKU (Fulfillment API returned empty sku + no lineItemProperties).
        if variation_by_txn:
            for order in orders:
                for li in order.line_items:
                    if li.sku or li.variation_specifics:
                        continue  # already resolved
                    item_id = li.legacy_item_id
                    txn_id = li.legacy_transaction_id
                    if not item_id or not txn_id:
                        continue
                    var_info = variation_by_txn.get((item_id, txn_id))
                    if var_info:
                        specs_str, var_sku = var_info
                        li.variation_specifics = specs_str
                        label = _variation_label(specs_str)
                        if label:
                            li.title = f"{li.title} [{label}]"
                        if var_sku:
                            li.sku = var_sku
                        log.debug(
                            "  variation from Trading API sold list: item=%s txn=%s → %r sku=%r",
                            item_id, txn_id, specs_str, var_sku,
                        )

        # GetOrders fallback: variation specifics for items still missing SKU/variation
        empty_sku_order_ids = list(dict.fromkeys(
            order.order_id
            for order in orders
            for li in order.line_items
            if not li.sku and not li.variation_specifics
            and li.legacy_item_id and li.legacy_transaction_id
        ))
        if empty_sku_order_ids:
            log.debug(
                "Variation data: GetOrders fallback for %d orders",
                len(empty_sku_order_ids),
            )
            extra_variations = self._fetch_order_variation_data(empty_sku_order_ids)
            for order in orders:
                for li in order.line_items:
                    if li.sku or li.variation_specifics:
                        continue
                    var_info = extra_variations.get(
                        (li.legacy_item_id, li.legacy_transaction_id)
                    )
                    if var_info:
                        specs_str, var_sku = var_info
                        li.variation_specifics = specs_str
                        label = _variation_label(specs_str)
                        if label:
                            li.title = f"{li.title} [{label}]"
                        if var_sku:
                            li.sku = var_sku
                        log.debug(
                            "  variation applied from GetOrders: item=%s txn=%s → %r sku=%r",
                            li.legacy_item_id, li.legacy_transaction_id, specs_str, var_sku,
                        )

    def _fetch_item_variations(self, item_id: str) -> list[tuple[dict[str, str], str]]:
        """
        Call Trading API GetItem and return [(variation_specifics_dict, sku), ...].
        variation_specifics_dict maps Name → Value (e.g. {'Size': 'Medium Light', 'String': '3G - .022'}).
        sku is the variation SKU, or "" if none is set on that variation.
        """
        xml_template = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetItemRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            f"<ItemID>{item_id}</ItemID>"
            "<IncludeVariations>true</IncludeVariations>"
            "<OutputSelector>Variations</OutputSelector>"
            "</GetItemRequest>"
        )
        ns = {"e": _TRADING_NS}
        result: list[tuple[dict[str, str], str]] = []
        try:
            root = self._call_trading_api(xml_template, "GetItem")
        except EbayAPIError as exc:
            log.warning("GetItem variations failed for item %s: %s", item_id, exc)
            return result
        for variation in root.findall(".//e:Variations/e:Variation", ns):
            sku_el = variation.find("e:SKU", ns)
            var_sku = (sku_el.text or "").strip() if sku_el is not None else ""
            var_specs: dict[str, str] = {}
            for nvl in variation.findall("e:VariationSpecifics/e:NameValueList", ns):
                name = _xml_text(nvl, "e:Name", ns)
                value = _xml_text(nvl, "e:Value", ns)
                if name:
                    var_specs[name] = value
            result.append((var_specs, var_sku))
        log.debug(
            "GetItem %s: %d variations found (%d with SKUs)",
            item_id, len(result), sum(1 for _, s in result if s),
        )
        return result

    def _enrich_variation_skus(self, orders: list[EbayOrder]) -> None:
        """
        Resolve variation specifics strings to real Neto SKUs.

        Priority order:
          1. CSV cache (EbayVariationSkuManager) — instant, no API call
          2. Trading API GetItem — fetches variation SKUs from the eBay listing;
             result is saved to the CSV for future fetches

        Line items where variation_specifics is empty are skipped.
        If a SKU is resolved it replaces li.sku; the title tag applied in
        _parse_order is preserved regardless.
        """
        if self._variation_sku_manager is None:
            return

        # Group by item_id so we only call GetItem once per listing
        needs_api: dict[str, list[EbayLineItem]] = {}

        for order in orders:
            for li in order.line_items:
                if not li.variation_specifics:
                    continue
                cached = self._variation_sku_manager.lookup(
                    li.legacy_item_id, li.variation_specifics
                )
                if cached is not None:
                    # cached == "" means "tried before, listing has no SKU set"
                    if cached:
                        li.sku = cached
                        log.debug(
                            "  variation SKU: item=%s %r → %s (CSV cache)",
                            li.legacy_item_id, li.variation_specifics, cached,
                        )
                    continue
                # Not yet cached — queue for API lookup
                needs_api.setdefault(li.legacy_item_id, []).append(li)

        if not needs_api:
            return

        log.debug("Variation SKU: API lookup needed for %d item IDs", len(needs_api))

        for item_id, line_items in needs_api.items():
            variations = self._fetch_item_variations(item_id)

            for li in line_items:
                our_specs = _parse_variation_specifics(li.variation_specifics)
                matched_sku = ""
                for var_specs, var_sku in variations:
                    if var_specs == our_specs:
                        matched_sku = var_sku
                        break

                if matched_sku:
                    li.sku = matched_sku
                    log.debug(
                        "  variation SKU: item=%s %r → %s (GetItem)",
                        item_id, li.variation_specifics, matched_sku,
                    )
                else:
                    log.debug(
                        "  variation SKU: item=%s %r → no SKU in listing",
                        item_id, li.variation_specifics,
                    )

                # Save to CSV (empty string if no SKU, to avoid future API calls)
                self._variation_sku_manager.save(
                    item_id, li.variation_specifics, matched_sku
                )

    def _get_txn_ids_for_orders(
        self,
        order_ids: list[str],
    ) -> dict[tuple[str, str], str]:
        """
        Use Trading API GetOrders to look up (order_id, item_id) → transaction_id.

        Called as a fallback when the Fulfillment API returns an empty
        legacyTransactionId for a listing that has sold multiple times, making
        it impossible to resolve the correct transaction from GetMyeBaySelling alone.

        Accepts new-format order IDs (e.g. "20-14397-93331") via OrderIDType=Platform.
        Returns {(order_id, item_id): txn_id} keyed by both the extended and legacy IDs
        so either format can be used for lookup.
        """
        if not order_ids:
            return {}

        ns = {"e": _TRADING_NS}
        order_ids_xml = "".join(f"<OrderID>{oid}</OrderID>" for oid in order_ids)
        xml_template = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetOrdersRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            f"<OrderIDArray>{order_ids_xml}</OrderIDArray>"
            "<OrderIDType>Platform</OrderIDType>"
            "</GetOrdersRequest>"
        )

        result: dict[tuple[str, str], str] = {}
        try:
            root = self._call_trading_api(xml_template, "GetOrders")
        except EbayAPIError as exc:
            log.warning("GetOrders txn ID resolution failed: %s", exc)
            return result

        for order_el in root.findall(".//e:Order", ns):
            ext_id = _xml_text(order_el, "e:ExtendedOrderID", ns)
            legacy_id = _xml_text(order_el, "e:OrderID", ns)
            for txn_el in order_el.findall(".//e:Transaction", ns):
                item_id = _xml_text(txn_el, "e:Item/e:ItemID", ns)
                txn_id_el = txn_el.find("e:TransactionID", ns)
                txn_id = (txn_id_el.text or "").strip() if txn_id_el is not None else ""
                if not item_id or not txn_id:
                    continue
                if ext_id:
                    result[(ext_id, item_id)] = txn_id
                if legacy_id:
                    result[(legacy_id, item_id)] = txn_id
                log.debug(
                    "  GetOrders: order=%s (ext=%s) item=%s → txn=%s",
                    legacy_id, ext_id, item_id, txn_id,
                )

        return result

    def _fetch_order_variation_data(
        self,
        order_ids: list[str],
    ) -> dict[tuple[str, str], tuple[str, str]]:
        """
        Use Trading API GetOrders to extract variation specifics per transaction.

        Returns {(item_id, txn_id): (specs_str, var_sku)}.
        """
        if not order_ids:
            return {}

        ns = {"e": _TRADING_NS}
        order_ids_xml = "".join(f"<OrderID>{oid}</OrderID>" for oid in order_ids)
        xml_template = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetOrdersRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            f"<OrderIDArray>{order_ids_xml}</OrderIDArray>"
            "<OrderIDType>Platform</OrderIDType>"
            "</GetOrdersRequest>"
        )

        result: dict[tuple[str, str], tuple[str, str]] = {}
        try:
            root = self._call_trading_api(xml_template, "GetOrders")
        except EbayAPIError as exc:
            log.warning("GetOrders variation data failed: %s", exc)
            return result

        for order_el in root.findall(".//e:Order", ns):
            for txn_el in order_el.findall(".//e:Transaction", ns):
                item_id = _xml_text(txn_el, "e:Item/e:ItemID", ns)
                txn_id_el = txn_el.find("e:TransactionID", ns)
                txn_id = (txn_id_el.text or "").strip() if txn_id_el is not None else ""
                if not item_id or not txn_id:
                    continue
                var_specs: dict[str, str] = {}
                for nvl in txn_el.findall("e:Variation/e:VariationSpecifics/e:NameValueList", ns):
                    name = _xml_text(nvl, "e:Name", ns)
                    value = _xml_text(nvl, "e:Value", ns)
                    if name:
                        var_specs[name] = value
                if not var_specs:
                    continue
                var_sku_el = txn_el.find("e:Variation/e:SKU", ns)
                var_sku = (var_sku_el.text or "").strip() if var_sku_el is not None else ""
                specs_str = " \u00b7".join(f"{k}: {v}" for k, v in var_specs.items())
                result[(item_id, txn_id)] = (specs_str, var_sku)
                log.debug(
                    "  GetOrders variation: item=%s txn=%s → %r sku=%r",
                    item_id, txn_id, specs_str, var_sku,
                )

        return result

    def set_private_notes(
        self,
        item_id: str,
        transaction_id: str,
        note_text: str,
        dry_run: bool = True,
    ) -> None:
        """
        Set (add or update) PrivateNotes on a sold eBay item via the Trading API.
        Uses SetUserNotes with ItemID + TransactionID to target a specific transaction.
        Note text is limited to 255 characters by eBay.
        """
        note_text = note_text[:255]

        if dry_run:
            print(f"[DRY RUN] eBay SetUserNotes on {item_id}-{transaction_id}: {note_text!r}")
            return

        if not self._config.dev_id:
            raise EbayAPIError(
                "Trading API credentials not configured (dev_id missing in config.json). "
                "Add your eBay Developer Program App ID (dev_id) to enable PrivateNotes."
            )

        log.debug(
            "SetUserNotes: item_id=%r transaction_id=%r note=%r",
            item_id, transaction_id, note_text,
        )
        xml_template = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<SetUserNotesRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            f"<ItemID>{item_id}</ItemID>"
            f"<TransactionID>{transaction_id}</TransactionID>"
            "<Action>AddOrUpdate</Action>"
            f"<NoteText>{_xml_escape(note_text)}</NoteText>"
            "</SetUserNotesRequest>"
        )
        self._call_trading_api(xml_template, "SetUserNotes")
        log.debug("SetUserNotes succeeded for item %s txn %s", item_id, transaction_id)

    def revise_item_shipping_dimensions(
        self,
        item_id: str,
        weight_kg: float,
        length_cm: float,
        width_cm: float,
        height_cm: float,
        dry_run: bool = True,
    ) -> None:
        """
        Update the shipping package dimensions on an active eBay listing via ReviseItem.
        Uses metric units (kg and cm) as required for eBay AU (site 15).

        weight_kg    — total package weight in kilograms
        length_cm    — longest dimension in centimetres
        width_cm     — second dimension in centimetres
        height_cm    — shortest dimension in centimetres
        """
        weight_major = int(weight_kg)                      # whole kg
        weight_minor = round((weight_kg - weight_major) * 1000)  # grams (0–999)

        if dry_run:
            print(
                f"[DRY RUN] eBay ReviseItem dimensions for ItemID {item_id}: "
                f"{weight_kg}kg  {length_cm}×{width_cm}×{height_cm}cm"
            )
            return

        if not self._config.dev_id:
            raise EbayAPIError(
                "Trading API credentials not configured (dev_id missing). "
                "Add dev_id to config.json to enable eBay dimension updates."
            )

        xml_template = (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<ReviseItemRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            "<Item>"
            f"<ItemID>{_xml_escape(item_id)}</ItemID>"
            "<ShippingPackageDetails>"
            f'<PackageDepth measurementSystem="Metric" unit="cm">{math.ceil(height_cm)}</PackageDepth>'
            f'<PackageLength measurementSystem="Metric" unit="cm">{math.ceil(length_cm)}</PackageLength>'
            f'<PackageWidth measurementSystem="Metric" unit="cm">{math.ceil(width_cm)}</PackageWidth>'
            f'<WeightMajor measurementSystem="Metric" unit="kg">{weight_major}</WeightMajor>'
            f'<WeightMinor measurementSystem="Metric" unit="gm">{weight_minor}</WeightMinor>'
            "</ShippingPackageDetails>"
            "</Item>"
            "</ReviseItemRequest>"
        )
        self._call_trading_api(xml_template, "ReviseItem")
        log.info(
            "eBay ReviseItem dimensions updated for ItemID %s: "
            "%skg  %sx%sx%scm",
            item_id, weight_kg, length_cm, width_cm, height_cm,
        )

    def get_item_images(self, legacy_item_ids: list[str]) -> dict[str, str]:
        """
        Fetch the primary listing image for each ItemID via Trading API GetItem.
        Returns {legacy_item_id: image_url}. Silently skips items that fail.
        """
        if not legacy_item_ids or not self._config.dev_id:
            return {}

        ns = {"e": _TRADING_NS}
        result = {}
        for item_id in legacy_item_ids:
            if not item_id:
                continue
            xml_template = (
                '<?xml version="1.0" encoding="utf-8"?>'
                f'<GetItemRequest xmlns="{_TRADING_NS}">'
                "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
                f"<ItemID>{item_id}</ItemID>"
                "<IncludeItemSpecifics>false</IncludeItemSpecifics>"
                "<OutputSelector>PictureDetails</OutputSelector>"
                "</GetItemRequest>"
            )
            try:
                root = self._call_trading_api(xml_template, "GetItem")
                pic_url = _xml_text(root, ".//e:PictureDetails/e:PictureURL", ns)
                if pic_url:
                    result[item_id] = pic_url
            except Exception:
                continue

        return result

    def _build_sold_list_xml(self, duration_days: int, page: int) -> str:
        """Build GetMyeBaySelling XML with {TOKEN} placeholder for _call_trading_api."""
        return (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetMyeBaySellingRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            "<SoldList>"
            "<Include>true</Include>"
            f"<DurationInDays>{duration_days}</DurationInDays>"
            "<IncludeNotes>true</IncludeNotes>"
            f"<Pagination><EntriesPerPage>200</EntriesPerPage><PageNumber>{page}</PageNumber></Pagination>"
            "</SoldList>"
            "</GetMyeBaySellingRequest>"
        )

    def _build_transactions_xml(self, date_from: datetime, date_to: datetime, page: int) -> str:
        """Build GetSellerTransactions XML with {TOKEN} placeholder for _call_trading_api."""
        from_str = date_from.strftime("%Y-%m-%dT00:00:00.000Z")
        to_str = date_to.strftime("%Y-%m-%dT23:59:59.000Z")
        return (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetSellerTransactionsRequest xmlns="{_TRADING_NS}">'
            "<RequesterCredentials><eBayAuthToken>{TOKEN}</eBayAuthToken></RequesterCredentials>"
            f"<CreateTimeFrom>{from_str}</CreateTimeFrom>"
            f"<CreateTimeTo>{to_str}</CreateTimeTo>"
            "<IncludeVariations>true</IncludeVariations>"
            f"<Pagination><EntriesPerPage>200</EntriesPerPage><PageNumber>{page}</PageNumber></Pagination>"
            "</GetSellerTransactionsRequest>"
        )

    # ----- Orders -----

    def _parse_order(self, raw: dict) -> EbayOrder:
        buyer = raw.get("buyer", {})
        buyer_name = (
            buyer.get("registrationAddress", {}).get("fullName", "")
            or buyer.get("username", "")
        )

        creation_raw = raw.get("creationDate", "")
        creation_date = _parse_ebay_date(creation_raw)

        line_items = []
        for li in raw.get("lineItems", []):
            image_url = ""
            image_data = li.get("image")
            if isinstance(image_data, dict):
                image_url = str(image_data.get("imageUrl", "") or "").strip()
            try:
                unit_price = float(str(li.get("lineItemCost", {}).get("value", 0) or 0))
            except (ValueError, TypeError):
                unit_price = 0.0

            sku_raw = str(li.get("sku", "") or "").strip()
            title = str(li.get("title", "")).strip()
            variation_specifics = ""

            if _is_variation_specifics(sku_raw):
                # Legacy path: eBay put variation specifics in the sku field
                # (middle-dot separated "Name: Value ·Name: Value").
                variation_specifics = sku_raw
                sku_raw = ""
                label = _variation_label(variation_specifics)
                if label:
                    title = f"{title} [{label}]"
            elif not sku_raw:
                # No SKU — check lineItemProperties for variation specifics.
                # eBay Fulfillment API returns [{name, value}, ...] here for
                # variation listings where the seller hasn't set per-variation SKUs.
                # Log all keys when SKU is empty so we can identify the right field.
                log.debug(
                    "  empty SKU on item=%s title=%r — raw li keys: %s",
                    li.get("legacyItemId", ""), li.get("title", "")[:40],
                    list(li.keys()),
                )
                props = li.get("lineItemProperties", [])
                if props:
                    parts = [
                        f"{p.get('name', '').strip()}: {p.get('value', '').strip()}"
                        for p in props
                        if p.get("name") and p.get("value")
                    ]
                    if parts:
                        variation_specifics = " \u00b7".join(parts)
                        label = _variation_label(variation_specifics)
                        if label:
                            title = f"{title} [{label}]"
                        log.debug(
                            "  variation specifics from lineItemProperties: item=%s %r",
                            li.get("legacyItemId", ""), variation_specifics,
                        )
                else:
                    log.debug(
                        "  no lineItemProperties for item=%s — checking other keys",
                        li.get("legacyItemId", ""),
                    )
                    # Log values of any keys that might contain variation data
                    for key in li:
                        val = li[key]
                        if isinstance(val, (list, dict)) and val:
                            log.debug("    li[%r] = %r", key, val)

            line_items.append(EbayLineItem(
                line_item_id=str(li.get("lineItemId", "")),
                sku=sku_raw,
                title=title,
                quantity=int(li.get("quantity", 1)),
                legacy_item_id=str(li.get("legacyItemId", "") or ""),
                legacy_transaction_id=str(li.get("legacyTransactionId", "") or ""),
                image_url=image_url,
                unit_price=unit_price,
                variation_specifics=variation_specifics,
            ))

        # Extract shipping address from fulfillmentStartInstructions
        ship_to = {}
        instructions = raw.get("fulfillmentStartInstructions", [])
        if instructions:
            ship_to = (
                instructions[0]
                .get("shippingStep", {})
                .get("shipTo", {})
            )
        contact = ship_to.get("contactAddress", {})

        # Pricing
        pricing = raw.get("pricingSummary", {})
        try:
            order_total = float(str(pricing.get("total", {}).get("value", 0) or 0))
        except (ValueError, TypeError):
            order_total = 0.0
        try:
            shipping_cost = float(str(pricing.get("deliveryCost", {}).get("value", 0) or 0))
        except (ValueError, TypeError):
            shipping_cost = 0.0

        # Shipping method from fulfillmentStartInstructions
        shipping_method = ""
        if instructions:
            shipping_method = str(
                instructions[0].get("shippingStep", {}).get("shippingServiceCode", "") or ""
            ).strip()

        shipping_type = _classify_ebay_shipping(shipping_method, instructions)

        # Tracking (populated once a shipping fulfillment has been created)
        fulfillments = raw.get("fulfillments") or []
        tracking_number = ""
        tracking_carrier = ""
        if fulfillments:
            first_f = fulfillments[0]
            tracking_number = str(first_f.get("shipmentTrackingNumber") or "").strip()
            tracking_carrier = str(first_f.get("shippingCarrierCode") or "").strip()

        return EbayOrder(
            order_id=str(raw.get("orderId", "")),
            buyer_name=buyer_name,
            buyer_notes=str(raw.get("buyerCheckoutNotes", "") or "").strip(),
            creation_date=creation_date,
            order_status=raw.get("orderFulfillmentStatus", ""),
            payment_status=raw.get("orderPaymentStatus", ""),
            line_items=line_items,
            ship_name=str(ship_to.get("fullName", "") or "").strip(),
            ship_street1=str(contact.get("addressLine1", "") or "").strip(),
            ship_street2=str(contact.get("addressLine2", "") or "").strip(),
            ship_city=str(contact.get("city", "") or "").strip(),
            ship_state=str(contact.get("stateOrProvince", "") or "").strip(),
            ship_postcode=str(contact.get("postalCode", "") or "").strip(),
            ship_country=str(contact.get("countryCode", "") or "").strip(),
            ship_phone=str(ship_to.get("primaryPhone", {}).get("phoneNumber", "") or "").strip(),
            order_total=order_total,
            shipping_cost=shipping_cost,
            shipping_method=shipping_method,
            shipping_type=shipping_type,
            tracking_number=tracking_number,
            tracking_carrier=tracking_carrier,
        )


def _classify_ebay_shipping(method: str, instructions: list[dict]) -> str:
    """Classify eBay shipping into Express, Regular, or Local Pickup."""
    # Check fulfillmentInstructionType for IN_STORE_PICKUP or SHIP_TO
    if instructions:
        instr_type = str(instructions[0].get("fulfillmentInstructionType", "") or "").upper()
        if "PICKUP" in instr_type:
            return "Local Pickup"

    m = method.lower()
    if not m:
        return ""
    if "pickup" in m or "collect" in m:
        return "Local Pickup"
    if "express" in m or "priority" in m or "overnight" in m or "next day" in m:
        return "Express"
    return "Regular"


def _xml_text(el: ET.Element, path: str, ns: dict) -> str:
    """Find an element by namespaced path and return its text, or empty string."""
    found = el.find(path, ns)
    return (found.text or "").strip() if found is not None else ""


def _extract_code(url_or_code: str) -> str | None:
    """Extract the 'code' query parameter from a URL, or return raw string as-is."""
    url_or_code = url_or_code.strip()
    if url_or_code.startswith("http"):
        parsed = urlparse(url_or_code)
        params = parse_qs(parsed.query)
        codes = params.get("code", [])
        return codes[0] if codes else None
    # Assume the user pasted the raw code
    return url_or_code or None


def _to_ebay_datetime(dt: datetime) -> str:
    """Convert a local datetime to eBay's ISO 8601 UTC format."""
    if dt.tzinfo is None:
        # .timestamp() treats naive datetimes as local time and converts correctly to UTC
        dt = datetime.fromtimestamp(dt.timestamp(), tz=timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_ebay_date(raw: str) -> datetime | None:
    if not raw:
        return None
    try:
        # eBay returns "2024-01-15T09:23:11.000Z"
        return datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return None


def _is_variation_specifics(sku: str) -> bool:
    """Return True if the SKU field from the Fulfillment API is actually a
    variation specifics string rather than a real SKU.
    eBay uses the middle dot (U+00B7) as a separator between specifics pairs.
    Note: newer eBay API responses use lineItemProperties instead; this handles
    the legacy case where specifics were packed into the sku field."""
    return "\u00b7" in sku


def _parse_variation_specifics(specifics: str) -> dict[str, str]:
    """Parse 'Size: Medium Light ·String: 3G - .022' → {'Size': 'Medium Light', 'String': '3G - .022'}."""
    result: dict[str, str] = {}
    for part in specifics.split("\u00b7"):
        part = part.strip()
        if ": " in part:
            name, _, value = part.partition(": ")
            result[name.strip()] = value.strip()
    return result


def _variation_label(specifics: str) -> str:
    """Extract just the values for display: 'Medium Light / 3G - .022'."""
    parsed = _parse_variation_specifics(specifics)
    return " / ".join(parsed.values()) if parsed else ""


def _xml_escape(text: str) -> str:
    """Escape special XML characters in user-provided text."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
