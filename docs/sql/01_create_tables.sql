-- =============================================================================
-- Scarlett AIO — Phase 1 + POS Core tables
-- Run this in the Supabase SQL editor (Project → SQL Editor → New query)
-- Safe to re-run: all statements use IF NOT EXISTS / OR REPLACE
-- =============================================================================

-- ---------------------------------------------------------------------------
-- SEQUENCES
-- ---------------------------------------------------------------------------

CREATE SEQUENCE IF NOT EXISTS transaction_number_seq START 1;
CREATE SEQUENCE IF NOT EXISTS quote_number_seq START 1;
CREATE SEQUENCE IF NOT EXISTS invoice_number_seq START 1;
CREATE SEQUENCE IF NOT EXISTS customer_id_seq START 1;


-- ---------------------------------------------------------------------------
-- 1. USERS
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS users (
    id              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    username        text UNIQUE NOT NULL,
    first_name      text NOT NULL,
    last_name       text NOT NULL,
    password_hash   text NOT NULL,
    role            text NOT NULL DEFAULT 'user',   -- 'admin' or 'user'
    is_active       boolean DEFAULT true,
    created_at      timestamptz DEFAULT now(),
    created_by      text,
    last_login_at   timestamptz
);


-- ---------------------------------------------------------------------------
-- 2. SUPPLIERS
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS suppliers (
    id                      text PRIMARY KEY,   -- short code e.g. 'DUNLOP'
    name                    text NOT NULL,
    abn                     text,
    account_number          text,
    payment_terms_days      integer DEFAULT 30,
    sku_suffix              text,
    sku_prefix              text,
    character_substitutions jsonb,
    address_line1           text,
    address_line2           text,
    city                    text,
    state                   text,
    postcode                text,
    notes                   text,
    active                  boolean DEFAULT true
);

CREATE TABLE IF NOT EXISTS supplier_contacts (
    id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    supplier_id text NOT NULL REFERENCES suppliers(id) ON DELETE CASCADE,
    role        text,
    name        text,
    email       text,
    phone       text,
    is_primary  boolean DEFAULT false
);


-- ---------------------------------------------------------------------------
-- 3. INVENTORY
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS items (
    id                              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    sku                             text UNIQUE NOT NULL,
    web_sku                         text,
    title                           text NOT NULL,
    brand                           text,
    series                          text,
    instrument                      text,
    sub_instrument                  text,
    supplier_id                     text REFERENCES suppliers(id),
    supplier_rrp                    numeric(10,2),
    last_purchase_cost              numeric(10,2),
    average_cost_exc_gst            numeric(10,2),
    average_cost_inc_gst            numeric(10,2),
    gst_amount                      numeric(10,2),
    minimum_sell                    numeric(10,2),
    online_sale_price               numeric(10,2),
    internal_barcode                text,
    product_barcode                 text,
    pick_zone                       text,
    qty_on_hand                     integer DEFAULT 0,
    qty_allocated_online            integer DEFAULT 0,
    qty_allocated_customer          integer DEFAULT 0,
    qty_on_order                    integer DEFAULT 0,
    min_order_level                 integer,
    max_order_level                 integer,
    is_kit                          boolean DEFAULT false,
    is_serialised                   boolean DEFAULT false,
    stock_availability_from_supplier boolean,
    description                     text,
    weight_kg                       numeric(8,3),
    length_cm                       numeric(8,2),
    width_cm                        numeric(8,2),
    height_cm                       numeric(8,2),
    last_purchase_date              date,
    last_sold_date                  date,
    created_date                    date DEFAULT CURRENT_DATE,
    active                          boolean DEFAULT true
);

-- Computed availability view
CREATE OR REPLACE VIEW item_availability AS
SELECT
    id,
    sku,
    title,
    qty_on_hand,
    qty_allocated_online,
    qty_allocated_customer,
    qty_on_hand - qty_allocated_online - qty_allocated_customer AS qty_available
FROM items;

CREATE TABLE IF NOT EXISTS stock_movements (
    id              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    item_id         uuid NOT NULL REFERENCES items(id),
    movement_type   text NOT NULL,
    qty_change      integer NOT NULL,
    reference_id    text,
    notes           text,
    performed_by    text,
    performed_at    timestamptz DEFAULT now()
);


-- ---------------------------------------------------------------------------
-- 5. CUSTOMERS & DISCOUNTS
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS discounts (
    id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    name        text NOT NULL,
    percentage  numeric(5,2) NOT NULL,
    is_system   boolean DEFAULT false,
    is_active   boolean DEFAULT true,
    created_at  timestamptz DEFAULT now()
);

-- Seed system presets (idempotent)
INSERT INTO discounts (name, percentage, is_system)
VALUES
    ('10%',  10.00, true),
    ('20%',  20.00, true),
    ('30%',  30.00, true),
    ('40%',  40.00, true),
    ('50%',  50.00, true)
ON CONFLICT DO NOTHING;

CREATE TABLE IF NOT EXISTS customers (
    id                      uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    customer_id             integer UNIQUE DEFAULT nextval('customer_id_seq'),
    customer_barcode        text UNIQUE,
    first_name              text NOT NULL,
    surname                 text,
    business                text,
    mobile                  text NOT NULL,
    phone_1                 text,
    fax                     text,
    email                   text,
    website                 text,
    address_1               text,
    address_2               text,
    city                    text,
    state                   text,
    postcode                text,
    country                 text DEFAULT 'Australia',
    ship_same_as_invoice    boolean DEFAULT true,
    ship_address_1          text,
    ship_address_2          text,
    ship_city               text,
    ship_state              text,
    ship_postcode           text,
    ship_country            text,
    tax_exemption_number    text,
    discount_id             uuid REFERENCES discounts(id),
    terms_days              integer,
    credit_limit            numeric(10,2),
    stop_credit             boolean DEFAULT false,
    is_local                boolean DEFAULT false,
    abn                     text,
    newsletter_opt_in       boolean DEFAULT false,
    private_comment         text,
    statement_comment       text,
    active                  boolean DEFAULT true,
    created_at              timestamptz DEFAULT now(),
    created_by              text,
    musipos_account_code    text,
    musipos_barcode_ref     text
);


-- ---------------------------------------------------------------------------
-- 6. POS TRANSACTIONS
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS transactions (
    id                      uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    transaction_number      text UNIQUE NOT NULL,
    quote_number            integer,
    invoice_number          integer,
    sale_type               text NOT NULL,              -- standard / quote / invoice / repair / deposit / refund
    sale_status             text NOT NULL DEFAULT 'completed',
    customer_id             uuid REFERENCES customers(id),
    staff_id                uuid REFERENCES users(id),  -- nullable; for future users migration
    performed_by            text,                        -- staff username from existing auth system
    subtotal                numeric(10,2) NOT NULL DEFAULT 0,
    cart_discount_pct       numeric(5,2),
    cart_discount_total     numeric(10,2),
    override_total          numeric(10,2),
    total                   numeric(10,2) NOT NULL DEFAULT 0,
    total_cost              numeric(10,2),
    payment_cash            numeric(10,2) DEFAULT 0,
    payment_eft             jsonb,
    payment_online          numeric(10,2) DEFAULT 0,
    cash_tendered           numeric(10,2),
    change_given            numeric(10,2),
    discount_id             uuid REFERENCES discounts(id),
    notes                   text,
    print_notes             boolean DEFAULT false,
    due_date                date,
    payment_terms_days      integer,
    linked_transaction_id   uuid REFERENCES transactions(id),
    park_name               text,
    cart_snapshot           jsonb,
    created_at              timestamptz DEFAULT now(),
    completed_at            timestamptz
);

CREATE TABLE IF NOT EXISTS transaction_lines (
    id              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    transaction_id  uuid NOT NULL REFERENCES transactions(id) ON DELETE CASCADE,
    item_id         uuid REFERENCES items(id),
    sku             text,
    description     text NOT NULL,
    qty             numeric(10,3) NOT NULL,
    unit_price      numeric(10,2) NOT NULL,
    cost_price      numeric(10,2),
    discount_pct    numeric(5,2) DEFAULT 0,
    line_total      numeric(10,2) NOT NULL,
    line_margin_pct numeric(5,2),
    is_refunded     boolean DEFAULT false,
    refunded_qty    numeric(10,3) DEFAULT 0
);


-- ---------------------------------------------------------------------------
-- DISABLE ROW LEVEL SECURITY (development — enable before production)
-- ---------------------------------------------------------------------------

ALTER TABLE users               DISABLE ROW LEVEL SECURITY;
ALTER TABLE suppliers           DISABLE ROW LEVEL SECURITY;
ALTER TABLE supplier_contacts   DISABLE ROW LEVEL SECURITY;
ALTER TABLE items               DISABLE ROW LEVEL SECURITY;
ALTER TABLE stock_movements     DISABLE ROW LEVEL SECURITY;
ALTER TABLE discounts           DISABLE ROW LEVEL SECURITY;
ALTER TABLE customers           DISABLE ROW LEVEL SECURITY;
ALTER TABLE transactions        DISABLE ROW LEVEL SECURITY;
ALTER TABLE transaction_lines   DISABLE ROW LEVEL SECURITY;


-- ---------------------------------------------------------------------------
-- INDEXES (basic performance)
-- ---------------------------------------------------------------------------

CREATE INDEX IF NOT EXISTS idx_items_sku            ON items(sku);
CREATE INDEX IF NOT EXISTS idx_items_web_sku        ON items(web_sku);
CREATE INDEX IF NOT EXISTS idx_items_barcode        ON items(internal_barcode);
CREATE INDEX IF NOT EXISTS idx_items_product_bc     ON items(product_barcode);
CREATE INDEX IF NOT EXISTS idx_transactions_num     ON transactions(transaction_number);
CREATE INDEX IF NOT EXISTS idx_transactions_cust    ON transactions(customer_id);
CREATE INDEX IF NOT EXISTS idx_transactions_status  ON transactions(sale_status);
CREATE INDEX IF NOT EXISTS idx_txn_lines_txn        ON transaction_lines(transaction_id);
CREATE INDEX IF NOT EXISTS idx_stock_mvmt_item      ON stock_movements(item_id);
CREATE INDEX IF NOT EXISTS idx_customers_mobile     ON customers(mobile);
CREATE INDEX IF NOT EXISTS idx_customers_barcode    ON customers(customer_barcode);
