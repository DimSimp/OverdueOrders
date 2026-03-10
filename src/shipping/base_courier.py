from __future__ import annotations

from abc import ABC, abstractmethod

from src.shipping.models import BookingResult, Quote, ShipmentRequest


class BaseCourier(ABC):
    """Abstract base class for all courier integrations."""

    name: str = ""   # Display name (e.g. "Australia Post")
    code: str = ""   # Internal ID for config lookup (e.g. "auspost")

    def __init__(self, config: dict):
        """
        Args:
            config: Courier-specific config dict from config.json shipping.couriers.<code>.
        """
        self._config = config

    @abstractmethod
    def get_quote(self, request: ShipmentRequest) -> list[Quote]:
        """
        Fetch shipping quote(s) for the given shipment.
        Returns a list of Quote objects (may include multiple service levels).
        On failure, return a single Quote with error set.
        """
        ...

    def book(self, request: ShipmentRequest, quote: Quote) -> BookingResult:
        """Book a shipment using a previously obtained quote. Phase 2 — not yet implemented."""
        return BookingResult(
            courier_name=self.name,
            tracking_number="",
            label_pdf=None,
            booking_reference="",
            error="Booking not yet implemented (Phase 2)",
        )

    def cancel_shipment(self, tracking_number: str, **kwargs) -> tuple[bool, str]:
        """Cancel a shipment by tracking number.

        Returns (success: bool, message: str).
        Override in subclasses that support cancellation.
        kwargs may contain courier-specific data (e.g. shipment_id for AusPost).
        """
        return False, f"Cancellation not supported for {self.name}"

    def is_available(self, request: ShipmentRequest) -> bool:
        """Check if this courier can handle the given shipment. Override for restrictions."""
        return True
