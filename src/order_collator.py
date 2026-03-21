from __future__ import annotations

import logging
import re
from collections import defaultdict
from dataclasses import dataclass

log = logging.getLogger(__name__)


_EBAY_PREFIX = re.compile(r'^ebay:[a-z0-9]+\s*', re.IGNORECASE)


def _normalize_street1(s: str) -> str:
    """Strip eBay-injected tracking prefix (e.g. 'ebay:tq5sqw7 ') from street1."""
    return _EBAY_PREFIX.sub('', s or '').strip().upper()


def _collation_key(order) -> tuple | None:
    """
    Return (platform, norm_street1, postcode, identity) or None if the order
    has insufficient address data for collation.

    identity = email for Neto orders, buyer_name for eBay orders.
    """
    is_neto = hasattr(order, 'date_placed')
    if is_neto:
        platform = order.sales_channel or 'Neto'
        street1 = _normalize_street1(getattr(order, 'ship_street1', ''))
        postcode = (getattr(order, 'ship_postcode', '') or '').strip().upper()
        identity = (order.email or '').strip().lower()
    else:
        platform = 'eBay'
        street1 = _normalize_street1(getattr(order, 'ship_street1', ''))
        postcode = (getattr(order, 'ship_postcode', '') or '').strip().upper()
        identity = (order.buyer_name or '').strip().lower()

    if is_neto:
        # Neto always has full address data — require all three fields
        if not street1 or not postcode or not identity:
            return None
    else:
        # eBay sometimes redacts street1 — postcode + buyer identity is sufficient
        if not postcode or not identity:
            return None
    return (platform, street1, postcode, identity)


@dataclass
class CollatedGroup:
    platform: str
    orders: list      # list of NetoOrder or EbayOrder (same platform)
    key: tuple        # (platform, street1, postcode, identity)

    @property
    def order_ids(self) -> list[str]:
        return [o.order_id for o in self.orders]

    @property
    def synthetic_id(self) -> str:
        """Unique fake order_id used as the treeview row key for this group."""
        return '__COLL__' + self.orders[0].order_id


def collate_orders(
    neto_orders: list,
    ebay_orders: list,
    ungrouped_ids: set,
) -> tuple[list, list, list]:
    """
    Detect orders going to the same address and group them.

    Returns:
        (collated_groups, remaining_neto, remaining_ebay)

    Orders whose order_id appears in *ungrouped_ids* are always treated as
    individual rows — they are never placed into a CollatedGroup.
    """

    def _group(orders):
        buckets: dict[tuple, list] = defaultdict(list)
        singles = []
        for o in orders:
            if o.order_id in ungrouped_ids:
                singles.append(o)
                continue
            key = _collation_key(o)
            log.debug("  collate key  %-24s  %s", o.order_id, key)
            if key is None:
                singles.append(o)
            else:
                buckets[key].append(o)
        groups = []
        for key, grp in buckets.items():
            if len(grp) >= 2:
                groups.append(CollatedGroup(platform=key[0], orders=grp, key=key))
            else:
                singles.extend(grp)
        return groups, singles

    neto_groups, neto_singles = _group(neto_orders)
    ebay_groups, ebay_singles = _group(ebay_orders)
    return neto_groups + ebay_groups, neto_singles, ebay_singles
