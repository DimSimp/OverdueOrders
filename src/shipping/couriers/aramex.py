from __future__ import annotations

import logging
import requests
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session

from src.shipping.base_courier import BaseCourier
from src.shipping.models import BookingResult, Quote, ShipmentRequest

log = logging.getLogger("courier.aramex")

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

    def _raw_address(self, request: ShipmentRequest) -> dict:
        """Build a plain address dict from the raw request (used as fallback)."""
        addr = request.receiver
        street = f"{addr.street1}, {addr.street2}".rstrip(", ")
        return {
            "streetAddress": street,
            "locality": addr.city,
            "stateOrProvince": addr.state,
            "postalCode": addr.postcode,
            "country": addr.country or "AU",
        }

    def _validate_address(self, request: ShipmentRequest) -> dict:
        """
        Validate receiver address via Aramex API.
        Returns the validated address dict on success, or the raw address dict
        as a fallback if validation is unavailable or returns an error.
        """
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

        try:
            resp = requests.post(
                f"{BASE_URL}/api/addresses/validate",
                headers=self._headers(),
                json=data,
                timeout=10,
            )
        except Exception as exc:
            log.warning("Address validation request failed (%s) — using raw address", exc)
            return self._raw_address(request)

        if resp.status_code != 200:
            log.warning("Address validation returned HTTP %d: %s — using raw address",
                        resp.status_code, resp.text[:400])
            return self._raw_address(request)

        result = resp.json()
        log.debug("Address validation response: %s", result)

        if "errors" in result:
            log.warning("Address validation errors: %s — using raw address", result["errors"])
            return self._raw_address(request)

        validated = result.get("data", {})
        if validated.get("stateOrProvince", "").lower() != addr.state.lower():
            log.warning("Address validation state mismatch (expected %s, got %s) — using raw address",
                        addr.state, validated.get("stateOrProvince"))
            return self._raw_address(request)

        log.debug("Address validated OK: %s", validated.get("streetAddress"))
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

        validated = self._validate_address(request)

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

    def book(self, request: ShipmentRequest, quote=None) -> BookingResult:
        """Create the consignment, return tracking number and label PDF bytes."""
        log.info("Booking Aramex consignment for order %s (receiver: %s, %s %s)",
                 request.order_id, request.receiver.name,
                 request.receiver.city, request.receiver.state)

        try:
            self._authenticate()
            log.debug("Aramex OAuth2 token obtained")
        except Exception as exc:
            log.error("Aramex auth failed: %s", exc)
            return BookingResult(courier_name=self.name, tracking_number="",
                                 label_pdf=None, booking_reference="", error=f"Auth failed: {exc}")

        log.debug("Validating receiver address: %s %s %s",
                  request.receiver.street1, request.receiver.city, request.receiver.postcode)
        validated = self._validate_address(request)

        signature = "SGR" if request.order_value > 200 else "ATL"
        log.debug("Delivery option: %s (order value $%.2f)", signature, request.order_value)

        items = []
        for pkg in request.packages:
            satchel = pkg.satchel_size
            if satchel:
                log.debug("Package: satchel %s", satchel)
                items.append({
                    "Quantity": 1,
                    "Reference": request.order_id,
                    "PackageType": "S",
                    "SatchelSize": satchel,
                })
            else:
                log.debug("Package: parcel %.2fkg  %dx%dx%dcm  cubic=%.3fkg",
                          pkg.weight_kg, pkg.length_cm, pkg.width_cm, pkg.height_cm,
                          pkg.cubic_weight_kg)
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

        consignment_data = {
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

        log.debug("POST %s/api/consignments  payload=%s", BASE_URL, consignment_data)
        try:
            resp = requests.post(
                f"{BASE_URL}/api/consignments",
                headers=self._headers(),
                json=consignment_data,
                timeout=20,
            )
            resp.raise_for_status()
            result = resp.json()
            log.debug("Consignment API response: %s", result)
            con_id = result["data"]["conId"]
            tracking_number = str(result["data"]["items"][0]["label"])
            log.info("Consignment created: conId=%s  tracking=%s", con_id, tracking_number)
        except Exception as exc:
            log.error("Aramex consignment POST failed: %s", exc)
            return BookingResult(courier_name=self.name, tracking_number="",
                                 label_pdf=None, booking_reference="", error=f"Booking failed: {exc}")

        # Download label PDF (4x6 format)
        label_pdf = None
        try:
            label_url = f"{BASE_URL}/api/consignments/{con_id}/labels?pageSize=4x6"
            log.debug("Downloading label PDF from %s", label_url)
            label_resp = requests.get(label_url, headers=self._headers(), timeout=20)
            label_resp.raise_for_status()
            label_pdf = label_resp.content
            log.info("Label PDF downloaded: %d bytes", len(label_pdf))
        except Exception as exc:
            log.warning("Label PDF download failed (non-fatal): %s", exc)

        return BookingResult(
            courier_name=self.name,
            tracking_number=tracking_number,
            label_pdf=label_pdf,
            booking_reference=str(con_id),
        )
