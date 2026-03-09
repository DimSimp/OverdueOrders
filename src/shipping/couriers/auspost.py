from __future__ import annotations

import requests
from requests.auth import HTTPBasicAuth

from src.shipping.base_courier import BaseCourier
from src.shipping.models import Quote, ShipmentRequest

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
