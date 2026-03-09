from __future__ import annotations

import xml.etree.ElementTree as ET

import requests

from src.shipping.base_courier import BaseCourier
from src.shipping.models import Quote, ShipmentRequest, next_business_day

QUOTE_URL = "https://appsrv.bondscouriers.com.au/bondsweb/api/upload-xml-job.htm"

# Service/vehicle combos to try (in order). First successful response wins.
SERVICE_VEHICLE_COMBOS = [
    ("C", "CAR"),
    ("C", "SW"),
    ("C", "SV"),
    ("TTK", ""),
]


class BondsCourier(BaseCourier):
    name = "Bonds Couriers"
    code = "bonds"

    def __init__(self, config: dict):
        super().__init__(config)
        self._account = config.get("account", "")
        self._auth_code = config.get("authorization_code", "")

    def is_available(self, request: ShipmentRequest) -> bool:
        return bool(self._account and self._auth_code)

    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        nbd = next_business_day()
        job_date = nbd.strftime("%Y-%m-%d")

        # Build dimension XML and totals
        dim_xml = ""
        total_items = 0
        total_weight = 0.0
        for pkg in request.packages:
            dims = sorted([pkg.length_cm, pkg.width_cm, pkg.height_cm])
            dim_xml += (
                f"<dimension>"
                f"<qty>1</qty>"
                f"<length>{dims[2]}</length>"
                f"<width>{dims[1]}</width>"
                f"<height>{dims[0]}</height>"
                f"</dimension>"
            )
            total_items += 1
            total_weight += pkg.weight_kg

        sender = request.sender
        receiver = request.receiver

        last_error = "All service/vehicle combos failed"

        for service_code, vehicle_code in SERVICE_VEHICLE_COMBOS:
            xml_payload = f'''
            <job xmlns:xi="http://www.w3.org/2001/XInclude"
                 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                 xsi:noNamespaceSchemaLocation="job-bonds.xsd">
            <job_action>QUOTE</job_action>
            <notifications>
            <notification>
            <notify_type>DELIVERY</notify_type>
            <notify_target>{receiver.email or "info@scarlettmusic.com.au"}</notify_target>
            </notification>
            <notification>
            <notify_type>DELIVERY</notify_type>
            <notify_target>{receiver.phone or ""}</notify_target>
            </notification>
            </notifications>
            <job_id/>
            <account>{self._account}</account>
            <authorization_code>{self._auth_code}</authorization_code>
            <containsDangerousGoods>false</containsDangerousGoods>
            <branch>MEL</branch>
            <job_date>{job_date}</job_date>
            <time_ready>09:00:00</time_ready>
            <deliver_by_time xsi:nil="true"/>
            <deliver_by_time_reason xsi:nil="true"/>
            <order_number>{request.order_id}</order_number>
            <contact>Kyal</contact>
            <insurance>true</insurance>
            <references>
            <reference>{request.order_id}</reference>
            </references>
            <service_code>{service_code}</service_code>
            <vehicle_code>{vehicle_code}</vehicle_code>
            <goods_description/>
            <instructions></instructions>
            <pallets/>
            <cubic/>
            <job_legs>
            <job_leg>
            <action>P</action>
            <service_qual/>
            <suburb>{sender.city}</suburb>
            <state>{sender.state}</state>
            <company>{sender.company or sender.name}</company>
            <address1>{sender.street1}</address1>
            <address2>{sender.street2}</address2>
            <contact>Kyal</contact>
            <items>{total_items}</items>
            <weight>{total_weight}</weight>
            <dimensions>
            {dim_xml}
            </dimensions>
            <references>
            <reference/>
            </references>
            </job_leg>
            <job_leg>
            <action>D</action>
            <service_qual></service_qual>
            <suburb>{receiver.city}</suburb>
            <state>{receiver.state}</state>
            <company>{receiver.company}</company>
            <address1>{receiver.street1}</address1>
            <address2>{receiver.street2}</address2>
            <contact/>
            <items>{total_items}</items>
            <weight>{total_weight}</weight>
            <dimensions>
            {dim_xml}
            </dimensions>
            <references>
            <reference/>
            </references>
            </job_leg>
            </job_legs>
            </job>
            '''

            try:
                resp = requests.post(
                    QUOTE_URL,
                    data=xml_payload,
                    headers={"Content-Type": "application/xml"},
                    timeout=10,
                )
                root = ET.fromstring(resp.text)
            except Exception as exc:
                last_error = str(exc)
                continue

            msg_status = root.findtext("msg_status", "")
            if msg_status == "ERROR":
                error_msg = root.findtext("msg_details", "Unknown error")
                last_error = f"{service_code}/{vehicle_code}: {error_msg}"
                continue

            # Parse price from job_details
            details = root.find("job_details")
            if details is None:
                last_error = f"{service_code}/{vehicle_code}: No job_details in response"
                continue

            try:
                fuel_charge = float(details.findtext("fuel_charge", "0"))
                job_charge = float(details.findtext("job_charge", "0"))
                gst = float(details.findtext("gst", "0"))
                price = fuel_charge + job_charge + gst
            except (ValueError, TypeError) as exc:
                last_error = f"{service_code}/{vehicle_code}: Price parse error: {exc}"
                continue

            if price <= 0:
                last_error = f"{service_code}/{vehicle_code}: Zero price returned"
                continue

            # Determine actual vehicle code from response for TTK
            actual_vehicle = vehicle_code
            if service_code != "C":
                actual_vehicle = details.findtext("vehicle_code", vehicle_code)

            return [Quote(
                courier_name=self.name,
                courier_code=self.code,
                service_name=f"Courier ({service_code}/{actual_vehicle})",
                price=round(price, 2),
                estimated_days="Same day / Next day",
                raw_response={
                    "service_code": service_code,
                    "vehicle_code": actual_vehicle,
                },
            )]

        # All combos failed
        return [Quote(
            courier_name=self.name, courier_code=self.code,
            service_name="", price=0, estimated_days="",
            error=last_error,
        )]
