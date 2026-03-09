from __future__ import annotations

import requests
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session

from src.shipping.base_courier import BaseCourier
from src.shipping.models import Quote, ShipmentRequest

TOKEN_URL = "https://identity.fastway.org/connect/token"
BASE_URL = "https://api.myfastway.com.au"
SCOPE = "fw-fl2-api-au"

# Cubic weight limit
MAX_CUBIC_WEIGHT_KG = 40


class AramexCourier(BaseCourier):
    name = "Aramex"
    code = "aramex"

    def __init__(self, config: dict):
        super().__init__(config)
        self._client_id = config.get("client_id", "")
        self._client_secret = config.get("client_secret", "")
        self._bearer_token: str | None = None

    def is_available(self, request: ShipmentRequest) -> bool:
        if not self._client_id or not self._client_secret:
            return False
        total_cubic = sum(pkg.cubic_weight_kg for pkg in request.packages)
        return total_cubic <= MAX_CUBIC_WEIGHT_KG

    def _authenticate(self) -> str:
        """Fetch OAuth2 bearer token from Fastway identity server."""
        client = BackendApplicationClient(client_id=self._client_id)
        oauth = OAuth2Session(client=client)
        token = oauth.fetch_token(
            token_url=TOKEN_URL,
            client_id=self._client_id,
            client_secret=self._client_secret,
            scope=SCOPE,
        )
        self._bearer_token = token["access_token"]
        return self._bearer_token

    def _headers(self) -> dict:
        token = self._bearer_token or self._authenticate()
        return {"Authorization": f"bearer {token}"}

    def _validate_address(self, request: ShipmentRequest) -> dict | None:
        """Validate receiver address. Returns validated address data or None on failure."""
        addr = request.receiver
        street = f"{addr.street1}, {addr.street2}".rstrip(", ")

        data = {
            "streetAddress": street,
            "additionalDetails": "",
            "locality": addr.city,
            "stateOrProvince": addr.state,
            "postalCode": addr.postcode,
            "country": addr.country or "AU",
            "lat": 0,
            "lng": 0,
            "userCreated": False,
        }

        resp = requests.post(
            f"{BASE_URL}/api/addresses/validate",
            headers=self._headers(),
            json=data,
            timeout=10,
        )
        result = resp.json()

        if "errors" in result:
            return None
        validated = result.get("data", {})
        if validated.get("stateOrProvince", "").lower() != addr.state.lower():
            return None
        return validated

    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        try:
            self._authenticate()
        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=f"Auth failed: {exc}",
            )]

        # Validate address
        try:
            validated = self._validate_address(request)
            if validated is None:
                return [Quote(
                    courier_name=self.name, courier_code=self.code,
                    service_name="", price=0, estimated_days="",
                    error="Address validation failed",
                )]
        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=f"Address validation error: {exc}",
            )]

        # Determine service code: ATL (authority to leave) for <$200, SGR (signature) for ≥$200
        signature = "SGR" if request.order_value > 200 else "ATL"

        # Build items array
        items = []
        for pkg in request.packages:
            satchel = pkg.satchel_size
            if satchel:
                items.append({
                    "Quantity": 1,
                    "Reference": request.order_id,
                    "PackageType": "S",
                    "SatchelSize": satchel,
                })
            else:
                items.append({
                    "Quantity": 1,
                    "Reference": request.order_id,
                    "PackageType": "P",
                    "WeightDead": pkg.weight_kg,
                    "WeightCubic": pkg.cubic_weight_kg,
                    "Length": pkg.length_cm,
                    "Width": pkg.width_cm,
                    "Height": pkg.height_cm,
                })

        quote_data = {
            "To": {
                "ContactName": request.receiver.name,
                "BusinessName": request.receiver.company,
                "PhoneNumber": request.receiver.phone or "0400000000",
                "Email": request.receiver.email or "info@scarlettmusic.com.au",
                "Address": {
                    "StreetAddress": validated.get("streetAddress", request.receiver.street1),
                    "Locality": validated.get("locality", request.receiver.city),
                    "StateOrProvince": validated.get("stateOrProvince", request.receiver.state),
                    "PostalCode": validated.get("postalCode", request.receiver.postcode),
                    "Country": validated.get("country", "AU"),
                },
            },
            "Services": [{"ServiceCode": "DELOPT", "ServiceItemCode": signature}],
            "Items": items,
            "ExternalRef1": request.order_id,
            "ExternalRef2": request.order_id,
        }

        try:
            resp = requests.post(
                f"{BASE_URL}/api/consignments/quote",
                headers=self._headers(),
                json=quote_data,
                timeout=15,
            )
            resp.raise_for_status()
            result = resp.json()
            price = result["data"]["total"]
            service_label = "Satchel" if any(i.get("PackageType") == "S" for i in items) else "Parcel"
            return [Quote(
                courier_name=self.name,
                courier_code=self.code,
                service_name=f"{service_label} ({signature})",
                price=float(price),
                estimated_days="1-3 business days",
                raw_response=result,
            )]
        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=str(exc),
            )]
