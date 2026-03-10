from __future__ import annotations

import logging

import requests
from requests.auth import HTTPBasicAuth

from src.shipping.base_courier import BaseCourier
from src.shipping.models import BookingResult, Quote, ShipmentRequest

log = logging.getLogger("courier.auspost")

QUOTE_URL = "https://digitalapi.auspost.com.au/shipping/v1/prices/shipments"

# Product IDs
STANDARD_PRODUCT = "3D55"
EXPRESS_PRODUCT = "3J55"

# Limits
MAX_DIMENSION_CM = 105
MAX_WEIGHT_KG = 22


class AusPostCourier(BaseCourier):
    name = "Australia Post"
    code = "auspost"

    def __init__(self, config: dict):
        super().__init__(config)
        self._account_number = config.get("account_number", "")
        self._username = config.get("api_username", "")
        self._secret = config.get("api_secret", "")

    def is_available(self, request: ShipmentRequest) -> bool:
        if not self._account_number or not self._username or not self._secret:
            return False
        for pkg in request.packages:
            if pkg.weight_kg > MAX_WEIGHT_KG:
                return False
            if any(d > MAX_DIMENSION_CM for d in (pkg.length_cm, pkg.width_cm, pkg.height_cm)):
                return False
        return True

    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        product_id = EXPRESS_PRODUCT if request.shipping_type == "Express" else STANDARD_PRODUCT
        service_label = "Express Post" if product_id == EXPRESS_PRODUCT else "Parcel Post"

        items = []
        for i, pkg in enumerate(request.packages, 1):
            items.append({
                "item_reference": f"{request.order_id}-{i}",
                "product_id": product_id,
                "length": pkg.length_cm,
                "height": pkg.height_cm,
                "width": pkg.width_cm,
                "weight": pkg.weight_kg,
                "authority_to_leave": "false",
                "allow_partial_delivery": "true",
            })

        payload = {
            "shipments": [{
                "from": {
                    "suburb": request.sender.city.upper(),
                    "state": request.sender.state,
                    "postcode": request.sender.postcode,
                },
                "to": {
                    "suburb": request.receiver.city.upper(),
                    "state": request.receiver.state,
                    "postcode": request.receiver.postcode,
                },
                "items": items,
            }]
        }

        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Account-Number": self._account_number,
        }

        try:
            resp = requests.post(
                QUOTE_URL,
                headers=headers,
                auth=HTTPBasicAuth(self._username, self._secret),
                json=payload,
                timeout=15,
            )
            resp.raise_for_status()
            data = resp.json()
            price = data["shipments"][0]["shipment_summary"]["total_cost"]
            return [Quote(
                courier_name=self.name,
                courier_code=self.code,
                service_name=service_label,
                price=float(price),
                estimated_days="2-6 business days" if product_id == STANDARD_PRODUCT else "1-3 business days",
                raw_response=data,
            )]
        except Exception as exc:
            return [Quote(
                courier_name=self.name,
                courier_code=self.code,
                service_name=service_label,
                price=0,
                estimated_days="",
                error=str(exc),
            )]

    def book(self, request: ShipmentRequest, quote=None) -> BookingResult:
        """Create an AusPost shipment, download the label PDF, return tracking + label."""
        product_id = EXPRESS_PRODUCT if request.shipping_type == "Express" else STANDARD_PRODUCT

        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Account-Number": self._account_number,
        }
        auth = HTTPBasicAuth(self._username, self._secret)

        # Build items array from packages
        items = []
        for i, pkg in enumerate(request.packages, 1):
            items.append({
                "item_reference": f"{request.order_id}-{i}",
                "product_id": product_id,
                "length": pkg.length_cm,
                "height": pkg.height_cm,
                "width": pkg.width_cm,
                "weight": pkg.weight_kg,
                "authority_to_leave": "false",
                "allow_partial_delivery": "true",
            })

        recv = request.receiver
        sender = request.sender
        shipment_payload = {
            "shipments": [{
                "shipment_reference": request.order_id,
                "customer_reference_1": request.order_id,
                "email_tracking_enabled": True,
                "from": {
                    "name": sender.name,
                    "business_name": sender.company,
                    "lines": [sender.street1],
                    "suburb": sender.city.upper(),
                    "state": sender.state,
                    "postcode": sender.postcode,
                    "phone": sender.phone,
                    "email": sender.email,
                },
                "to": {
                    "name": recv.name,
                    "business_name": (recv.company or recv.name)[:40],
                    "lines": [recv.street1[:50], (recv.street2 or "")[:50]],
                    "suburb": recv.city.upper(),
                    "state": recv.state,
                    "postcode": recv.postcode,
                    "phone": recv.phone or "",
                    "email": recv.email or "",
                },
                "items": items,
            }]
        }

        log.info("Booking AusPost shipment for order %s (product=%s)", request.order_id, product_id)
        log.debug("AusPost shipment payload: %s", shipment_payload)

        try:
            resp = requests.post(
                "https://digitalapi.auspost.com.au/shipping/v1/shipments",
                headers=headers,
                auth=auth,
                json=shipment_payload,
                timeout=20,
            )
            log.debug("AusPost shipment HTTP %d: %s", resp.status_code, resp.text[:500])
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="", error=f"AusPost shipment creation failed: {exc}",
            )

        if "errors" in data:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="", error=f"AusPost API error: {data['errors']}",
            )

        try:
            shipment = data["shipments"][0]
            shipment_id = shipment["shipment_id"]
            item_ids = [{"item_id": x["item_id"]} for x in shipment["items"]]
            tracking = shipment["items"][0]["tracking_details"]["consignment_id"]
        except (KeyError, IndexError) as exc:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="", error=f"AusPost response missing fields: {exc}",
            )

        log.info("AusPost shipment created: tracking=%s  shipment_id=%s", tracking, shipment_id)

        # Request label PDF
        label_payload = {
            "wait_for_label_url": True,
            "preferences": [{
                "type": "PRINT",
                "format": "PDF",
                "groups": [
                    {
                        "group": "Parcel Post",
                        "layout": "THERMAL-LABEL-A6-1PP",
                        "branded": True,
                        "left_offset": 0,
                        "top_offset": 0,
                    },
                    {
                        "group": "Express Post",
                        "layout": "THERMAL-LABEL-A6-1PP",
                        "branded": False,
                        "left_offset": 0,
                        "top_offset": 0,
                    },
                ],
            }],
            "shipments": [{"shipment_id": shipment_id, "items": item_ids}],
        }

        try:
            resp = requests.post(
                "https://digitalapi.auspost.com.au/shipping/v1/labels",
                headers=headers,
                auth=auth,
                json=label_payload,
                timeout=20,
            )
            log.debug("AusPost label HTTP %d: %s", resp.status_code, resp.text[:300])
            resp.raise_for_status()
            label_data = resp.json()
            label_url = label_data["labels"][0]["url"]
        except Exception as exc:
            log.error("AusPost label request failed: %s", exc)
            return BookingResult(
                courier_name=self.name,
                tracking_number=str(tracking),
                label_pdf=None,
                booking_reference=str(shipment_id),
                error=f"Shipment created (tracking: {tracking}) but label request failed: {exc}",
            )

        try:
            pdf_resp = requests.get(label_url, timeout=30)
            pdf_resp.raise_for_status()
            label_pdf = pdf_resp.content
            log.info("AusPost label downloaded: %d bytes  url=%s", len(label_pdf), label_url)
        except Exception as exc:
            log.error("AusPost label download failed: %s", exc)
            return BookingResult(
                courier_name=self.name,
                tracking_number=str(tracking),
                label_pdf=None,
                booking_reference=str(shipment_id),
                error=f"Shipment created (tracking: {tracking}) but label download failed: {exc}",
            )

        return BookingResult(
            courier_name=self.name,
            tracking_number=str(tracking),
            label_pdf=label_pdf,
            booking_reference=str(shipment_id),
        )

    def cancel_shipment(self, tracking_number: str, **kwargs) -> tuple[bool, str]:
        """Cancel an AusPost shipment.

        Requires the shipment_id (passed as kwarg or looked up via the API).
        """
        shipment_id = kwargs.get("shipment_id", "")

        if not shipment_id:
            # Try to find the shipment by listing recent shipments
            shipment_id = self._find_shipment_id(tracking_number)

        if not shipment_id:
            return False, (
                f"Could not find shipment ID for tracking {tracking_number}. "
                "AusPost requires the shipment ID to cancel."
            )

        headers = {
            "Accept": "application/json",
            "Account-Number": self._account_number,
        }
        auth = HTTPBasicAuth(self._username, self._secret)

        log.info("Cancelling AusPost shipment: shipment_id=%s  tracking=%s",
                 shipment_id, tracking_number)

        try:
            resp = requests.delete(
                f"https://digitalapi.auspost.com.au/shipping/v1/shipments/{shipment_id}",
                headers=headers,
                auth=auth,
                timeout=20,
            )
            log.debug("AusPost cancel HTTP %d: %s", resp.status_code, resp.text[:500])

            if resp.status_code in (200, 202, 204):
                return True, f"Shipment {shipment_id} cancelled successfully."
            resp.raise_for_status()
            return True, f"Shipment {shipment_id} cancelled (HTTP {resp.status_code})."
        except Exception as exc:
            log.error("AusPost cancel failed: %s", exc)
            return False, str(exc)

    def _find_shipment_id(self, tracking_number: str) -> str:
        """Look up a shipment ID from tracking number via the AusPost shipments endpoint."""
        headers = {
            "Accept": "application/json",
            "Account-Number": self._account_number,
        }
        auth = HTTPBasicAuth(self._username, self._secret)

        try:
            resp = requests.get(
                "https://digitalapi.auspost.com.au/shipping/v1/shipments",
                headers=headers,
                auth=auth,
                timeout=15,
            )
            resp.raise_for_status()
            data = resp.json()
            for shipment in data.get("shipments", []):
                for item in shipment.get("items", []):
                    tracking = item.get("tracking_details", {}).get("consignment_id", "")
                    if tracking == tracking_number:
                        return shipment.get("shipment_id", "")
        except Exception as exc:
            log.warning("AusPost shipment lookup failed: %s", exc)
        return ""
