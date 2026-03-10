from __future__ import annotations

import base64
import logging
import re
from datetime import datetime

from src.shipping.base_courier import BaseCourier
from src.shipping.models import BookingResult, Quote, ShipmentRequest, next_business_day

log = logging.getLogger("courier.allied")

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

    # ── Helpers ───────────────────────────────────────────────────────────

    def _create_client(self, with_history: bool = False):
        """Create a zeep SOAP client.  Returns (client, history_plugin | None)."""
        import zeep
        from zeep.transports import Transport
        from zeep.plugins import HistoryPlugin

        history = HistoryPlugin() if with_history else None
        plugins = [history] if history else []
        transport = Transport(timeout=15, operation_timeout=15)
        client = zeep.Client(wsdl=WSDL_URL, transport=transport, plugins=plugins)
        client.transport.session.proxies = {"http": PROXY_URL}
        return client, history

    def _build_job(self, request: ShipmentRequest, account):
        """Build the Allied Express job dict from a ShipmentRequest."""
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

        # Extract numeric job number from order ID, capped to Java int range
        job_number_str = re.sub(r"[^0-9]", "", request.order_id)
        job_number = int(job_number_str) % 2_000_000_000 if job_number_str else 0

        nd = next_business_day()
        pickup_date = nd.replace(hour=10, minute=0, second=0, microsecond=0)
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
            "readyDate": pickup_date,
            "jobNumber": job_number,
            "vehicle": {"vehicleID": 1},
        }
        return job, job_number

    # ── Quote ─────────────────────────────────────────────────────────────

    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        try:
            import zeep  # noqa: F401
        except ImportError:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error="zeep library not installed (pip install zeep)",
            )]

        try:
            client, _ = self._create_client()
            account = client.service.getAccountDefaults(
                self._api_key, self._account_code, self._state, "AOE"
            )
        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=f"SOAP connection failed: {exc}",
            )]

        job, _ = self._build_job(request, account)

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

    # ── Booking ───────────────────────────────────────────────────────────

    def book(self, request: ShipmentRequest, quote=None) -> BookingResult:
        """Book an Allied Express shipment: validate → save → dispatch → get label."""
        try:
            import zeep  # noqa: F401
            import xmltodict
            from lxml import etree
        except ImportError as exc:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="",
                error=f"Missing dependency: {exc}. Install with: pip install zeep xmltodict lxml",
            )

        log.info("Booking Allied Express shipment for order %s", request.order_id)

        try:
            client, history = self._create_client(with_history=True)
            account = client.service.getAccountDefaults(
                self._api_key, self._account_code, self._state, "AOE"
            )
        except Exception as exc:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="", error=f"SOAP connection failed: {exc}",
            )

        job, job_number = self._build_job(request, account)
        job_ids = {"jobIds": job_number}

        # Validate
        try:
            job = client.service.validateBooking(self._api_key, job)
        except Exception as exc:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="", error=f"Validation failed: {exc}",
            )

        # Save + dispatch
        try:
            with client.settings(strict=False):
                client.service.savePendingJob(self._api_key, job)
                client.service.dispatchPendingJobs(self._api_key, job_ids)
                xml_str = etree.tostring(
                    history.last_received["envelope"], encoding="unicode"
                )
                xml_dict = xmltodict.parse(xml_str)
        except Exception as exc:
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="", error=f"Dispatch failed: {exc}",
            )

        # Extract connote (tracking) number from dispatch response
        try:
            dispatch_result = (
                xml_dict["soapenv:Envelope"]["soapenv:Body"]
                ["ns1:dispatchPendingJobsResponse"]["result"]["item"]
            )
            connote_number = dispatch_result["docketNumber"]
            reference = dispatch_result["referenceNumbers"]
        except (KeyError, TypeError) as exc:
            log.error("Could not parse dispatch response: %s\n%s", exc, xml_dict)
            return BookingResult(
                courier_name=self.name, tracking_number="", label_pdf=None,
                booking_reference="",
                error=f"Dispatch succeeded but could not parse response: {exc}",
            )

        log.info("Allied dispatch OK: connote=%s  reference=%s", connote_number, reference)

        # Get label PDF
        try:
            with client.settings(strict=False):
                client.service.getLabel(
                    self._api_key, "AOE", connote_number, reference,
                    request.sender.postcode, 1,
                )
                label_xml_str = etree.tostring(
                    history.last_received["envelope"], encoding="unicode"
                )
                label_xml = xmltodict.parse(label_xml_str)
            label_b64 = (
                label_xml["soapenv:Envelope"]["soapenv:Body"]
                ["ns1:getLabelResponse"]["result"]
            )
            label_pdf = base64.b64decode(label_b64)
            log.info("Allied label downloaded: %d bytes", len(label_pdf))
        except Exception as exc:
            log.error("Allied label download failed: %s", exc)
            return BookingResult(
                courier_name=self.name,
                tracking_number=str(connote_number),
                label_pdf=None,
                booking_reference=str(reference),
                error=f"Booking confirmed (tracking: {connote_number}) but label failed: {exc}",
            )

        return BookingResult(
            courier_name=self.name,
            tracking_number=str(connote_number),
            label_pdf=label_pdf,
            booking_reference=str(reference),
        )

    # ── Cancellation ──────────────────────────────────────────────────────

    def cancel_shipment(self, tracking_number: str, **kwargs) -> tuple[bool, str]:
        """Cancel an Allied Express shipment via SOAP cancelDispatchJob.

        Requires the destination postcode (passed as kwarg).
        """
        postcode = kwargs.get("postcode", "")
        if not postcode:
            return False, (
                "Destination postcode is required to cancel an Allied Express shipment. "
                "Please use the manual entry with a booking from today's list."
            )

        try:
            import zeep  # noqa: F401
        except ImportError:
            return False, "zeep library not installed (pip install zeep)"

        log.info("Cancelling Allied Express shipment: tracking=%s  postcode=%s",
                 tracking_number, postcode)

        try:
            client, _ = self._create_client()

            result = client.service.cancelDispatchJob(
                self._api_key, tracking_number, postcode
            )
            result_str = str(result)
            log.info("Allied cancel response: %s", result_str)

            if result_str == "0":
                return True, "Shipment cancelled successfully."
            elif result_str == "-6":
                return True, "Shipment was already cancelled."
            else:
                return False, f"Unexpected response from Allied: {result_str}"
        except Exception as exc:
            log.error("Allied cancel failed: %s", exc)
            return False, str(exc)
