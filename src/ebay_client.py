from __future__ import annotations

import base64
import time
import webbrowser
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timezone
from urllib.parse import urlencode, parse_qs, urlparse

import requests

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


class EbayAuthError(Exception):
    pass


class EbayAPIError(Exception):
    pass


class EbayClient:
    def __init__(self, config: EbayConfig, token_save_callback):
        """
        token_save_callback(access_token, expires_at, refresh_token=None)
        called after any token exchange or refresh to persist tokens.
        """
        self._config = config
        self._save_tokens = token_save_callback
        self._session = requests.Session()

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

            if not batch or len(all_raw) >= total:
                break
            offset += limit

        # Filter to PAID orders (eBay fulfillment filter doesn't support payment status)
        paid = [o for o in all_raw if o.get("orderPaymentStatus") == "PAID"]
        orders = [self._parse_order(o) for o in paid]

        # Enrich buyer_notes with PrivateNotes from the Trading API (if credentials configured)
        if self._config.dev_id:
            self._enrich_with_private_notes(orders, date_from, date_to)

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

    def _enrich_with_private_notes(
        self,
        orders: list[EbayOrder],
        date_from: datetime,
        date_to: datetime,
    ) -> None:
        """
        Fetch PrivateNotes from the Trading API (GetMyeBaySelling) and populate
        buyer_notes on each EbayOrder. PrivateNotes are returned per OrderLineItemID
        (ItemID-TransactionID). We extract the ItemID prefix and match against
        legacyItemId from the Fulfillment API.

        If any line item on an order has a PrivateNote, the combined notes are
        appended to buyer_notes so filter_on_po can detect them.
        """
        token = self._ensure_valid_token()

        # GetMyeBaySelling uses DurationInDays, not a date range.
        # Calculate days since date_from; cap at 60 (API limit for sold list).
        from datetime import timezone as _tz
        now = datetime.now()
        duration_days = min(int((now - date_from).days) + 1, 60)

        # Build lookup: ItemID (extracted from OLI prefix) → PrivateNotes text
        notes_by_item_id: dict[str, str] = {}
        page = 1
        while True:
            xml_body = self._build_sold_list_xml(token, duration_days, page)
            try:
                resp = self._session.post(
                    self._trading_url,
                    headers={
                        "X-EBAY-API-CALL-NAME": "GetMyeBaySelling",
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
            except Exception:
                break  # Trading API unavailable — silently skip note enrichment

            root = ET.fromstring(resp.text)
            ns = {"e": _TRADING_NS}

            ack = root.find("e:Ack", ns)
            if ack is None or ack.text not in ("Success", "Warning"):
                break

            for txn in root.findall(".//e:Transaction", ns):
                oli_el = txn.find("e:OrderLineItemID", ns)
                if oli_el is None or not oli_el.text:
                    continue
                # OLI format: "ItemID-TransactionID" — extract ItemID prefix
                item_id = oli_el.text.split("-")[0]
                note = _xml_text(txn, "e:Item/e:PrivateNotes", ns)
                if item_id and note:
                    notes_by_item_id[item_id] = note

            # Paginate through SoldList pages
            total_pages_el = root.find(
                ".//e:SoldList/e:PaginationResult/e:TotalNumberOfPages", ns
            )
            total_pages = int(total_pages_el.text) if total_pages_el is not None and total_pages_el.text else 1
            if page >= total_pages:
                break
            page += 1

        if not notes_by_item_id:
            return

        # Apply notes per line item; also roll up into order.buyer_notes for filter_on_po
        for order in orders:
            item_notes = []
            for li in order.line_items:
                note = notes_by_item_id.get(li.legacy_item_id, "")
                li.notes = note
                if note:
                    item_notes.append(note)
            if item_notes:
                combined = " | ".join(item_notes)
                if order.buyer_notes:
                    order.buyer_notes = order.buyer_notes + " | " + combined
                else:
                    order.buyer_notes = combined

    def get_item_images(self, legacy_item_ids: list[str]) -> dict[str, str]:
        """
        Fetch the primary listing image for each ItemID via Trading API GetItem.
        Returns {legacy_item_id: image_url}. Silently skips items that fail.
        """
        if not legacy_item_ids:
            return {}
        try:
            token = self._ensure_valid_token()
        except Exception:
            return {}

        result = {}
        for item_id in legacy_item_ids:
            if not item_id:
                continue
            xml_body = (
                '<?xml version="1.0" encoding="utf-8"?>'
                f'<GetItemRequest xmlns="{_TRADING_NS}">'
                f"<RequesterCredentials><eBayAuthToken>{token}</eBayAuthToken></RequesterCredentials>"
                f"<ItemID>{item_id}</ItemID>"
                "<IncludeItemSpecifics>false</IncludeItemSpecifics>"
                "<OutputSelector>PictureDetails</OutputSelector>"
                "</GetItemRequest>"
            )
            try:
                resp = self._session.post(
                    self._trading_url,
                    headers={
                        "X-EBAY-API-CALL-NAME": "GetItem",
                        "X-EBAY-API-APP-NAME": self._config.client_id,
                        "X-EBAY-API-DEV-NAME": self._config.dev_id,
                        "X-EBAY-API-CERT-NAME": self._config.client_secret,
                        "X-EBAY-API-SITEID": EBAY_AU_SITE_ID,
                        "X-EBAY-API-COMPATIBILITY-LEVEL": EBAY_TRADING_VERSION,
                        "Content-Type": "text/xml",
                    },
                    data=xml_body.encode("utf-8"),
                    timeout=15,
                )
                resp.raise_for_status()
                root = ET.fromstring(resp.text)
                ns = {"e": _TRADING_NS}
                ack = root.find("e:Ack", ns)
                if ack is not None and ack.text in ("Success", "Warning"):
                    pic_url = _xml_text(root, ".//e:PictureDetails/e:PictureURL", ns)
                    if pic_url:
                        result[item_id] = pic_url
            except Exception:
                continue

        return result

    def _build_sold_list_xml(self, token: str, duration_days: int, page: int) -> str:
        return (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetMyeBaySellingRequest xmlns="{_TRADING_NS}">'
            f"<RequesterCredentials><eBayAuthToken>{token}</eBayAuthToken></RequesterCredentials>"
            "<SoldList>"
            "<Include>true</Include>"
            f"<DurationInDays>{duration_days}</DurationInDays>"
            "<IncludeNotes>true</IncludeNotes>"
            f"<Pagination><EntriesPerPage>200</EntriesPerPage><PageNumber>{page}</PageNumber></Pagination>"
            "</SoldList>"
            "</GetMyeBaySellingRequest>"
        )

    def _build_transactions_xml(self, token: str, date_from: datetime, date_to: datetime, page: int) -> str:
        from_str = date_from.strftime("%Y-%m-%dT00:00:00.000Z")
        to_str = date_to.strftime("%Y-%m-%dT23:59:59.000Z")
        return (
            '<?xml version="1.0" encoding="utf-8"?>'
            f'<GetSellerTransactionsRequest xmlns="{_TRADING_NS}">'
            f"<RequesterCredentials><eBayAuthToken>{token}</eBayAuthToken></RequesterCredentials>"
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
            line_items.append(EbayLineItem(
                line_item_id=str(li.get("lineItemId", "")),
                sku=str(li.get("sku", "") or "").strip(),
                title=str(li.get("title", "")).strip(),
                quantity=int(li.get("quantity", 1)),
                legacy_item_id=str(li.get("legacyItemId", "") or ""),
                legacy_transaction_id=str(li.get("legacyTransactionId", "") or ""),
                image_url=image_url,
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
        )


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
