from __future__ import annotations

import base64
import time
import webbrowser
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
EBAY_SCOPE = "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly"


@dataclass
class EbayLineItem:
    line_item_id: str
    sku: str
    title: str
    quantity: int


@dataclass
class EbayOrder:
    order_id: str
    buyer_name: str
    buyer_notes: str
    creation_date: datetime | None
    order_status: str       # orderFulfillmentStatus
    payment_status: str     # orderPaymentStatus
    line_items: list[EbayLineItem] = field(default_factory=list)


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
        return [self._parse_order(o) for o in paid]

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
            line_items.append(EbayLineItem(
                line_item_id=str(li.get("lineItemId", "")),
                sku=str(li.get("sku", "") or "").strip(),
                title=str(li.get("title", "")).strip(),
                quantity=int(li.get("quantity", 1)),
            ))

        return EbayOrder(
            order_id=str(raw.get("orderId", "")),
            buyer_name=buyer_name,
            buyer_notes=str(raw.get("buyerCheckoutNotes", "") or "").strip(),
            creation_date=creation_date,
            order_status=raw.get("orderFulfillmentStatus", ""),
            payment_status=raw.get("orderPaymentStatus", ""),
            line_items=line_items,
        )


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
