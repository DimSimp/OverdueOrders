from __future__ import annotations

import re
from datetime import datetime

from src.shipping.base_courier import BaseCourier
from src.shipping.models import Quote, ShipmentRequest

WSDL_URL = "http://neptune.alliedexpress.com.au:8080/ttws-ejb/TTWS?wsdl"
PROXY_URL = "http://neptune.alliedexpress.com.au:8080/ttws-ejb/TTWS"

# Markup: 26.9% margin + 10% GST (from legacy code)
MARKUP_FACTOR = 1.269 * 1.1

SERVICE_LEVEL = "R"  # Road Express / Overnight


class AlliedCourier(BaseCourier):
    name = "Allied Express"
    code = "allied"

    def __init__(self, config: dict):
        super().__init__(config)
        self._api_key = config.get("api_key", "")
        self._account_code = config.get("account_code", "")
        self._state = config.get("state", "VIC")

    def is_available(self, request: ShipmentRequest) -> bool:
        if not self._api_key or not self._account_code:
            return False
        try:
            import zeep  # noqa: F401
            return True
        except ImportError:
            return False

    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        try:
            import zeep
            from zeep.transports import Transport
            from zeep.plugins import HistoryPlugin
        except ImportError:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error="zeep library not installed (pip install zeep)",
            )]

        try:
            history = HistoryPlugin()
            transport = Transport(timeout=15, operation_timeout=15)
            client = zeep.Client(wsdl=WSDL_URL, transport=transport, plugins=[history])
            client.transport.session.proxies = {"http": PROXY_URL}

            # Get account defaults
            account = client.service.getAccountDefaults(
                self._api_key, self._account_code, self._state, "AOE"
            )
        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=f"SOAP connection failed: {exc}",
            )]

        sender = request.sender
        receiver = request.receiver

        pickup_address = {
            "address1": sender.street1,
            "address2": sender.street2,
            "country": "Australia",
            "postCode": sender.postcode,
            "state": sender.state,
            "suburb": sender.city,
        }

        delivery_address = {
            "address1": receiver.street1,
            "address2": receiver.street2,
            "country": "Australia",
            "postCode": receiver.postcode,
            "state": receiver.state,
            "suburb": receiver.city,
        }

        pickup_stop = {
            "companyName": sender.company or sender.name,
            "contact": "Kyal Scarlett",
            "emailAddress": sender.email or "info@scarlettmusic.com.au",
            "geographicAddress": pickup_address,
            "phoneNumber": sender.phone or "03 9318 5751",
            "stopNumber": 1,
            "stopType": "P",
        }

        delivery_stop = {
            "companyName": receiver.company or receiver.name,
            "contact": receiver.name,
            "emailAddress": receiver.email,
            "geographicAddress": delivery_address,
            "phoneNumber": receiver.phone,
            "stopNumber": 2,
            "stopType": "D",
        }

        # Build items array
        cubed_items = []
        total_volume = 0.0
        total_weight = 0.0
        total_items = 0
        total_cubic = 0.0

        for pkg in request.packages:
            volume = pkg.volume_m3
            cubed_items.append({
                "dangerous": "false",
                "height": pkg.height_cm,
                "itemCount": 1,
                "length": pkg.length_cm,
                "volume": volume,
                "weight": pkg.weight_kg,
                "width": pkg.width_cm,
            })
            total_volume += volume
            total_weight += pkg.weight_kg
            total_items += 1
            total_cubic += pkg.cubic_weight_kg

        # Extract numeric job number from order ID
        job_number_str = re.sub(r"[^0-9]", "", request.order_id)
        job_number = int(job_number_str) if job_number_str else 0

        today = datetime.now().strftime("%Y-%m-%d")
        pickup_instructions = "The music shop, open 9am-6pm. Best parking is at The Palms across the road."

        job = {
            "account": account,
            "cubicWeight": total_cubic,
            "Docket": "SCM",
            "instructions": pickup_instructions,
            "cubedItems": cubed_items,
            "itemCount": total_items,
            "weight": total_weight,
            "volume": total_volume,
            "items": cubed_items,
            "jobStops": [pickup_stop, delivery_stop],
            "serviceLevel": SERVICE_LEVEL,
            "referenceNumbers": request.order_id,
            "bookedBy": "Kyal Scarlett",
            "readyDate": today,
            "jobNumber": job_number,
            "vehicle": {"vehicleID": 1},
        }

        try:
            job = client.service.validateBooking(self._api_key, job)
            job_price = client.service.calculatePrice(self._api_key, job)
            raw_price = float(job_price["totalCharge"])
            if raw_price <= 0:
                return [Quote(
                    courier_name=self.name, courier_code=self.code,
                    service_name="Road Express", price=0, estimated_days="",
                    error="Zero price returned",
                )]
            price = round(raw_price * MARKUP_FACTOR, 2)
            return [Quote(
                courier_name=self.name,
                courier_code=self.code,
                service_name="Road Express",
                price=price,
                estimated_days="Overnight / Next day",
                raw_response={"totalCharge": raw_price},
            )]
        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="Road Express", price=0, estimated_days="",
                error=str(exc),
            )]
