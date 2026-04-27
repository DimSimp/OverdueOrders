"""
Cross-platform dispatch syncing between eBay and Neto.

When an eBay order is dispatched directly, the corresponding Neto order (identified
by its PurchaseOrderNumber field, which Neto auto-populates with the eBay order ID)
should also be marked Dispatched.

When a Neto eBay-channel order is dispatched, the corresponding eBay order should
also be marked fulfilled via the eBay Fulfillment API.

Both functions return (success: bool, message: str) and are designed to be called
after the primary dispatch succeeds.  Failures are non-fatal — the primary dispatch
is already done.
"""
from __future__ import annotations

import logging
from datetime import datetime, timedelta
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from src.neto_client import NetoClient, NetoOrder
    from src.ebay_client import EbayClient

log = logging.getLogger(__name__)


def sync_ebay_to_neto(
    ebay_order_id: str,
    neto_client: "NetoClient",
    tracking_number: str = "",
    shipping_method: str = "",
    ebay_created_at: datetime | None = None,
    dry_run: bool = True,
) -> tuple[bool, str]:
    """After dispatching an eBay order, find the matching Neto order by
    PurchaseOrderNumber and mark it Dispatched.

    Returns (success, human-readable message).
    """
    try:
        date_from = date_to = None
        if ebay_created_at is not None:
            date_from = ebay_created_at - timedelta(days=1)
            date_to = ebay_created_at + timedelta(days=1)
        orders = neto_client.get_order_by_purchase_order_number(
            ebay_order_id,
            date_from=date_from,
            date_to=date_to,
        )
    except Exception as exc:
        log.warning("cross_dispatch: Neto PO# lookup failed for %s: %s", ebay_order_id, exc)
        return False, f"Neto lookup failed: {exc}"

    if not orders:
        log.debug("cross_dispatch: no Neto order found for PO# %s", ebay_order_id)
        return False, f"No Neto order found for PO# {ebay_order_id}"

    neto_order = orders[0]
    if neto_order.status == "Dispatched":
        log.debug("cross_dispatch: Neto order %s already Dispatched", neto_order.order_id)
        return True, f"Neto order {neto_order.order_id} already Dispatched"

    try:
        skus = [li.sku for li in neto_order.line_items]
        neto_client.update_order_status(
            neto_order.order_id,
            new_status="Dispatched",
            tracking_number=tracking_number,
            shipping_method=shipping_method,
            line_item_skus=skus,
            dry_run=dry_run,
        )
        prefix = "[DRY RUN] " if dry_run else ""
        msg = f"{prefix}Neto order {neto_order.order_id} marked Dispatched"
        log.info("cross_dispatch: %s", msg)
        return True, msg
    except Exception as exc:
        log.warning("cross_dispatch: Neto update failed for %s: %s", neto_order.order_id, exc)
        return False, f"Neto update failed: {exc}"


def sync_neto_to_ebay(
    neto_order: "NetoOrder",
    ebay_client: "EbayClient",
    tracking_number: str = "",
    carrier: str = "",
    dry_run: bool = True,
) -> tuple[bool, str]:
    """After dispatching a Neto eBay-channel order, find and fulfill the matching
    eBay order using the PurchaseOrderNumber (= eBay order ID).

    Returns (success, human-readable message).
    """
    ebay_order_id = (neto_order.purchase_order_number or "").strip()
    if not ebay_order_id:
        return False, "No eBay order ID in Neto PO# field"

    if neto_order.sales_channel.lower() != "ebay":
        return False, "Not an eBay-channel Neto order"

    try:
        ebay_orders = ebay_client.get_order_by_exact_id(ebay_order_id)
    except Exception as exc:
        log.warning("cross_dispatch: eBay lookup failed for %s: %s", ebay_order_id, exc)
        return False, f"eBay lookup failed: {exc}"

    if not ebay_orders:
        log.debug("cross_dispatch: no eBay order found for ID %s", ebay_order_id)
        return False, f"No eBay order found for ID {ebay_order_id}"

    ebay_order = ebay_orders[0]
    if ebay_order.order_status == "FULFILLED":
        log.debug("cross_dispatch: eBay order %s already FULFILLED", ebay_order_id)
        return True, f"eBay order {ebay_order_id} already fulfilled"

    try:
        ebay_client.create_shipping_fulfillment(
            ebay_order_id,
            line_items=ebay_order.line_items,
            tracking_number=tracking_number,
            carrier=carrier,
            dry_run=dry_run,
        )
        prefix = "[DRY RUN] " if dry_run else ""
        msg = f"{prefix}eBay order {ebay_order_id} marked fulfilled"
        log.info("cross_dispatch: %s", msg)
        return True, msg
    except Exception as exc:
        log.warning("cross_dispatch: eBay fulfillment failed for %s: %s", ebay_order_id, exc)
        return False, f"eBay fulfillment failed: {exc}"
