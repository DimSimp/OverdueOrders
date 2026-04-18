-- Migration 02: Transaction number generation + atomic stock adjustment
-- Run in Supabase SQL Editor after 01_create_tables.sql has been applied.
-- This script is idempotent — safe to re-run.

-- ── 1. Counter table ────────────────────────────────────────────────────────
-- One row per calendar year; last_seq is incremented atomically for each sale.

CREATE TABLE IF NOT EXISTS transaction_counters (
    year     integer PRIMARY KEY,
    last_seq integer NOT NULL DEFAULT 0
);

-- ── 2. Atomic number generator ──────────────────────────────────────────────
-- INSERT ... ON CONFLICT DO UPDATE ... RETURNING is a single atomic operation
-- in PostgreSQL, so two concurrent tills cannot receive the same number.
-- Numbers auto-reset each new year (a new row is created for the new year).

CREATE OR REPLACE FUNCTION next_transaction_number()
RETURNS text
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
    yr       integer;
    next_seq integer;
BEGIN
    yr := EXTRACT(YEAR FROM now())::integer;

    INSERT INTO transaction_counters (year, last_seq)
    VALUES (yr, 1)
    ON CONFLICT (year) DO UPDATE
        SET last_seq = transaction_counters.last_seq + 1
    RETURNING last_seq INTO next_seq;

    -- Produces e.g. 'T-2026-0001', 'T-2026-0002', ...
    RETURN 'T-' || yr || '-' || LPAD(next_seq::text, 4, '0');
END;
$$;

-- ── 3. Trigger: auto-assign on INSERT ───────────────────────────────────────
-- The Python app inserts a transaction without supplying transaction_number;
-- the trigger fills it in automatically before the row is written.

CREATE OR REPLACE FUNCTION trg_set_transaction_number()
RETURNS trigger
LANGUAGE plpgsql
AS $$
BEGIN
    IF NEW.transaction_number IS NULL OR NEW.transaction_number = '' THEN
        NEW.transaction_number := next_transaction_number();
    END IF;
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS set_transaction_number ON transactions;
CREATE TRIGGER set_transaction_number
    BEFORE INSERT ON transactions
    FOR EACH ROW EXECUTE FUNCTION trg_set_transaction_number();

-- ── 4. Atomic stock adjustment RPC ──────────────────────────────────────────
-- Called by the POS app via supabase.rpc('adjust_item_qty', {...}).
-- Using a server-side function makes the UPDATE atomic at the row level even
-- when multiple tills write concurrently. Pass a negative p_qty_change for
-- sales (stock out) and a positive value for refunds/receives (stock in).

CREATE OR REPLACE FUNCTION adjust_item_qty(p_item_id uuid, p_qty_change numeric)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
BEGIN
    UPDATE items
    SET qty_on_hand = qty_on_hand + p_qty_change
    WHERE id = p_item_id;
END;
$$;
