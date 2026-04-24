-- =============================================================================
-- Customer profile field additions for the current POS customer-profile work.
-- Safe to re-run.
-- =============================================================================

ALTER TABLE customers
    ADD COLUMN IF NOT EXISTS ship_same_as_invoice boolean DEFAULT true,
    ADD COLUMN IF NOT EXISTS ship_address_1 text,
    ADD COLUMN IF NOT EXISTS ship_address_2 text,
    ADD COLUMN IF NOT EXISTS ship_city text,
    ADD COLUMN IF NOT EXISTS ship_state text,
    ADD COLUMN IF NOT EXISTS ship_postcode text,
    ADD COLUMN IF NOT EXISTS ship_country text,
    ADD COLUMN IF NOT EXISTS discount_profile text;

ALTER TABLE customers
    ALTER COLUMN mobile DROP NOT NULL;
