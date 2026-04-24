-- =============================================================================
-- Migration 04 — add cash_rounding column to transactions
-- Run in Supabase SQL Editor.  Safe to re-run (IF NOT EXISTS / idempotent).
-- =============================================================================

ALTER TABLE transactions
    ADD COLUMN IF NOT EXISTS cash_rounding numeric(6,2);
