-- ⚠️  DEV RESET — DO NOT RUN IN PRODUCTION ⚠️
-- Wipes all transaction data and restores stock levels to pre-testing state.
-- Safe to run multiple times (idempotent).
-- Run in Supabase SQL Editor only during development.

-- ── Step 1: Reverse qty_on_hand for all test sales ──────────────────────────
-- stock_movements.qty_change is negative for sales, so subtracting it adds
-- the quantity back (double-negative = positive).

UPDATE items i
SET qty_on_hand = i.qty_on_hand - sm.total_sold
FROM (
    SELECT item_id, SUM(qty_change) AS total_sold
    FROM stock_movements
    WHERE movement_type = 'sale_instore'
    GROUP BY item_id
) sm
WHERE i.id = sm.item_id;

-- ── Step 2: Remove sale stock movement records ───────────────────────────────
-- Only removes sale_instore movements — leaves receive/stocktake/etc. intact.

DELETE FROM stock_movements
WHERE movement_type = 'sale_instore';

-- ── Step 3: Remove all transactions ─────────────────────────────────────────
-- transaction_lines is deleted automatically via ON DELETE CASCADE.

DELETE FROM transactions;

-- ── Step 4: Reset transaction number counter ─────────────────────────────────
-- Next sale after reset will be T-YYYY-0001.

TRUNCATE transaction_counters;
