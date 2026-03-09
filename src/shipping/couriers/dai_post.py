from __future__ import annotations

from pathlib import Path

from src.shipping.base_courier import BaseCourier
from src.shipping.models import Quote, ShipmentRequest

MAX_WEIGHT_KG = 22
MAX_ORDER_VALUE = 200

# Weight brackets: (max_weight, column_letter)
# Columns B-K in the rates sheet
WEIGHT_BRACKETS = [
    (0.5, "B"),
    (1.0, "C"),
    (1.5, "D"),
    (2.0, "D"),
    (2.5, "E"),
    (3.0, "E"),
    (3.5, "F"),
    (4.0, "F"),
    (4.5, "G"),
    (5.0, "G"),
    (7.0, "H"),
    (10.0, "I"),
    (15.0, "J"),
    (22.0, "K"),
]

# Map column letters to 0-based column indices
_COL_INDEX = {chr(c): c - ord("A") for c in range(ord("A"), ord("L") + 1)}


class DaiPostCourier(BaseCourier):
    name = "DAI Post"
    code = "dai_post"

    def __init__(self, config: dict):
        super().__init__(config)
        self._postcodes_file = config.get("postcodes_file", "")
        self._rates_file = config.get("rates_file", "")

    def is_available(self, request: ShipmentRequest) -> bool:
        if not self._postcodes_file or not self._rates_file:
            return False
        # Single parcel only
        if len(request.packages) > 1:
            return False
        # Max $200 order value
        if request.order_value > MAX_ORDER_VALUE:
            return False
        # Max weight
        if request.packages and request.packages[0].weight_kg > MAX_WEIGHT_KG:
            return False
        # Express not supported
        if request.shipping_type == "Express":
            return False
        return True

    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        try:
            import openpyxl
        except ImportError:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error="openpyxl library not installed",
            )]

        postcodes_path = Path(self._postcodes_file)
        rates_path = Path(self._rates_file)

        if not postcodes_path.exists():
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=f"Postcodes file not found: {postcodes_path}",
            )]
        if not rates_path.exists():
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=f"Rates file not found: {rates_path}",
            )]

        pkg = request.packages[0]
        weight = round(pkg.weight_kg, 2)
        postcode = request.receiver.postcode.strip()

        try:
            pc_wb = openpyxl.load_workbook(str(postcodes_path), read_only=True, data_only=True)
            pc_sheet = pc_wb[pc_wb.sheetnames[0]]

            # Find zone for postcode
            zone = None
            for row in pc_sheet.iter_rows(min_row=2, max_col=2, values_only=True):
                sheet_pc = str(row[0] or "").strip()
                if len(sheet_pc) == 3:
                    sheet_pc = f"0{sheet_pc}"
                if sheet_pc == postcode:
                    zone = str(row[1] or "").strip()
                    break
            pc_wb.close()

            if not zone:
                return [Quote(
                    courier_name=self.name, courier_code=self.code,
                    service_name="", price=0, estimated_days="",
                    error=f"Postcode {postcode} not found in DAI Post zones",
                )]

            # Find rate for zone + weight
            rates_wb = openpyxl.load_workbook(str(rates_path), read_only=True, data_only=True)
            rates_sheet = rates_wb[rates_wb.sheetnames[0]]

            # Determine weight column
            col_letter = None
            for max_w, col in WEIGHT_BRACKETS:
                if weight <= max_w:
                    col_letter = col
                    break

            if col_letter is None:
                rates_wb.close()
                return [Quote(
                    courier_name=self.name, courier_code=self.code,
                    service_name="", price=0, estimated_days="",
                    error=f"Weight {weight}kg exceeds DAI Post limit",
                )]

            col_idx = _COL_INDEX[col_letter]

            price = None
            for row in rates_sheet.iter_rows(min_row=2, values_only=True):
                rate_zone = str(row[0] or "").strip()
                if rate_zone.lower() == zone.lower():
                    try:
                        price = float(row[col_idx]) * 1.1  # +10% GST
                    except (ValueError, TypeError, IndexError):
                        pass
                    break
            rates_wb.close()

            if price is None:
                return [Quote(
                    courier_name=self.name, courier_code=self.code,
                    service_name="", price=0, estimated_days="",
                    error=f"No rate found for zone '{zone}' at {weight}kg",
                )]

            # PO Box surcharge
            addr = (request.receiver.street1 + " " + request.receiver.street2).lower()
            if "po box" in addr or "p.o. box" in addr:
                price += 0.80

            return [Quote(
                courier_name=self.name,
                courier_code=self.code,
                service_name="Standard Parcel",
                price=round(price, 2),
                estimated_days="3-7 business days",
                raw_response={"zone": zone, "weight": weight},
            )]

        except Exception as exc:
            return [Quote(
                courier_name=self.name, courier_code=self.code,
                service_name="", price=0, estimated_days="",
                error=str(exc),
            )]
