from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError, as_completed
from typing import Callable

from src.shipping.base_courier import BaseCourier
from src.shipping.models import Quote, ShipmentRequest

QUOTE_TIMEOUT = 30  # seconds per courier


class QuoteEngine:
    def __init__(self, couriers: list[BaseCourier]):
        self._couriers = couriers

    def get_quotes(
        self,
        request: ShipmentRequest,
        enabled_codes: set[str] | None = None,
        progress_callback: Callable[[str, str], None] | None = None,
    ) -> list[Quote]:
        """
        Fetch quotes from all enabled/available couriers in parallel.

        Args:
            request: The shipment details.
            enabled_codes: If set, only query couriers whose code is in this set.
            progress_callback: Called as (courier_name, status) where status is
                               "quoting", "done", or "error".

        Returns:
            List of Quote objects, sorted by price (successful first, errors last).
        """
        active = []
        for c in self._couriers:
            if enabled_codes is not None and c.code not in enabled_codes:
                continue
            if not c.is_available(request):
                continue
            active.append(c)

        if not active:
            return []

        all_quotes: list[Quote] = []

        def _fetch(courier: BaseCourier) -> list[Quote]:
            if progress_callback:
                progress_callback(courier.name, "quoting")
            try:
                quotes = courier.get_quote(request)
                status = "error" if all(q.error for q in quotes) else "done"
                if progress_callback:
                    progress_callback(courier.name, status)
                return quotes
            except Exception as exc:
                if progress_callback:
                    progress_callback(courier.name, "error")
                return [Quote(
                    courier_name=courier.name,
                    courier_code=courier.code,
                    service_name="",
                    price=0,
                    estimated_days="",
                    error=str(exc),
                )]

        executor = ThreadPoolExecutor(max_workers=len(active))
        futures = {executor.submit(_fetch, c): c for c in active}
        try:
            for future in as_completed(futures, timeout=QUOTE_TIMEOUT):
                try:
                    quotes = future.result(timeout=QUOTE_TIMEOUT)
                    all_quotes.extend(quotes)
                except Exception as exc:
                    courier = futures[future]
                    all_quotes.append(Quote(
                        courier_name=courier.name,
                        courier_code=courier.code,
                        service_name="",
                        price=0,
                        estimated_days="",
                        error=f"Timeout/error: {exc}",
                    ))
        except FuturesTimeoutError:
            for future, courier in futures.items():
                if not future.done():
                    all_quotes.append(Quote(
                        courier_name=courier.name,
                        courier_code=courier.code,
                        service_name="",
                        price=0,
                        estimated_days="",
                        error=f"Timed out (no response within {QUOTE_TIMEOUT}s)",
                    ))
        finally:
            executor.shutdown(wait=False)

        # Sort: successful quotes by price ascending, errors at the end
        successful = sorted([q for q in all_quotes if not q.error], key=lambda q: q.price)
        failed = [q for q in all_quotes if q.error]
        return successful + failed
