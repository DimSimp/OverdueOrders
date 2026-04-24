-- =============================================================================
-- Migration 05 - customer merge audit + atomic merge RPC
-- Run in Supabase SQL Editor after 01-04 have been applied.
-- Safe to re-run.
-- =============================================================================

CREATE TABLE IF NOT EXISTS customer_merge_audit (
    id                          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    merged_customer_id          uuid REFERENCES customers(id) ON DELETE SET NULL,
    source_customer_a_id        uuid NOT NULL,
    source_customer_b_id        uuid NOT NULL,
    source_customer_a_snapshot  jsonb NOT NULL,
    source_customer_b_snapshot  jsonb NOT NULL,
    selected_values             jsonb NOT NULL DEFAULT '{}'::jsonb,
    merged_customer_snapshot    jsonb,
    moved_transaction_count     integer NOT NULL DEFAULT 0,
    moved_parked_count          integer NOT NULL DEFAULT 0,
    merged_by                   text,
    merged_at                   timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_customer_merge_audit_merged_customer
    ON customer_merge_audit(merged_customer_id);

CREATE INDEX IF NOT EXISTS idx_customer_merge_audit_source_a
    ON customer_merge_audit(source_customer_a_id);

CREATE INDEX IF NOT EXISTS idx_customer_merge_audit_source_b
    ON customer_merge_audit(source_customer_b_id);


CREATE OR REPLACE FUNCTION merge_customers_fn(
    p_customer_a uuid,
    p_customer_b uuid,
    p_selected jsonb DEFAULT '{}'::jsonb,
    p_merged_by text DEFAULT NULL
)
RETURNS TABLE (
    audit_id uuid,
    merged_customer_uuid uuid,
    merged_customer_number integer,
    moved_transaction_count integer,
    moved_parked_count integer
)
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
    a customers%ROWTYPE;
    b customers%ROWTYPE;
    v_selected jsonb := COALESCE(p_selected, '{}'::jsonb);
    v_new_customer_uuid uuid;
    v_new_customer_number integer;
    v_new_customer_name text;
    v_merged_customer_snapshot jsonb;
    v_audit_id uuid;
    v_tx_count integer := 0;
    v_parked_count integer := 0;
BEGIN
    IF p_customer_a IS NULL OR p_customer_b IS NULL THEN
        RAISE EXCEPTION 'Two customer IDs are required for a merge.';
    END IF;

    IF p_customer_a = p_customer_b THEN
        RAISE EXCEPTION 'A customer profile cannot be merged with itself.';
    END IF;

    SELECT * INTO a
    FROM customers
    WHERE id = p_customer_a;

    IF NOT FOUND THEN
        RAISE EXCEPTION 'Customer A (%) was not found.', p_customer_a;
    END IF;

    SELECT * INTO b
    FROM customers
    WHERE id = p_customer_b;

    IF NOT FOUND THEN
        RAISE EXCEPTION 'Customer B (%) was not found.', p_customer_b;
    END IF;

    INSERT INTO customers (
        first_name,
        surname,
        business,
        mobile,
        phone_1,
        fax,
        email,
        website,
        address_1,
        address_2,
        city,
        state,
        postcode,
        country,
        ship_same_as_invoice,
        ship_address_1,
        ship_address_2,
        ship_city,
        ship_state,
        ship_postcode,
        ship_country,
        tax_exemption_number,
        discount_id,
        discount_profile,
        terms_days,
        credit_limit,
        stop_credit,
        is_local,
        abn,
        newsletter_opt_in,
        private_comment,
        statement_comment,
        active,
        created_by,
        musipos_account_code,
        musipos_barcode_ref
    )
    VALUES (
        COALESCE(NULLIF(BTRIM(v_selected->>'first_name'), ''), a.first_name, b.first_name),
        COALESCE(NULLIF(BTRIM(v_selected->>'surname'), ''), a.surname, b.surname),
        COALESCE(NULLIF(BTRIM(v_selected->>'business'), ''), a.business, b.business),
        COALESCE(NULLIF(BTRIM(v_selected->>'mobile'), ''), a.mobile, b.mobile),
        COALESCE(NULLIF(BTRIM(v_selected->>'phone_1'), ''), a.phone_1, b.phone_1),
        COALESCE(NULLIF(BTRIM(v_selected->>'fax'), ''), a.fax, b.fax),
        COALESCE(NULLIF(BTRIM(v_selected->>'email'), ''), a.email, b.email),
        COALESCE(NULLIF(BTRIM(v_selected->>'website'), ''), a.website, b.website),
        COALESCE(NULLIF(BTRIM(v_selected->>'address_1'), ''), a.address_1, b.address_1),
        COALESCE(NULLIF(BTRIM(v_selected->>'address_2'), ''), a.address_2, b.address_2),
        COALESCE(NULLIF(BTRIM(v_selected->>'city'), ''), a.city, b.city),
        COALESCE(NULLIF(BTRIM(v_selected->>'state'), ''), a.state, b.state),
        COALESCE(NULLIF(BTRIM(v_selected->>'postcode'), ''), a.postcode, b.postcode),
        COALESCE(NULLIF(BTRIM(v_selected->>'country'), ''), a.country, b.country, 'Australia'),
        CASE
            WHEN v_selected ? 'ship_same_as_invoice'
                AND jsonb_typeof(v_selected->'ship_same_as_invoice') <> 'null'
            THEN (v_selected->>'ship_same_as_invoice')::boolean
            ELSE COALESCE(a.ship_same_as_invoice, b.ship_same_as_invoice, true)
        END,
        COALESCE(NULLIF(BTRIM(v_selected->>'ship_address_1'), ''), a.ship_address_1, b.ship_address_1),
        COALESCE(NULLIF(BTRIM(v_selected->>'ship_address_2'), ''), a.ship_address_2, b.ship_address_2),
        COALESCE(NULLIF(BTRIM(v_selected->>'ship_city'), ''), a.ship_city, b.ship_city),
        COALESCE(NULLIF(BTRIM(v_selected->>'ship_state'), ''), a.ship_state, b.ship_state),
        COALESCE(NULLIF(BTRIM(v_selected->>'ship_postcode'), ''), a.ship_postcode, b.ship_postcode),
        COALESCE(NULLIF(BTRIM(v_selected->>'ship_country'), ''), a.ship_country, b.ship_country, 'Australia'),
        COALESCE(
            NULLIF(BTRIM(v_selected->>'tax_exemption_number'), ''),
            a.tax_exemption_number,
            b.tax_exemption_number
        ),
        CASE
            WHEN v_selected ? 'discount_id'
                AND NULLIF(BTRIM(v_selected->>'discount_id'), '') IS NOT NULL
            THEN (v_selected->>'discount_id')::uuid
            ELSE COALESCE(a.discount_id, b.discount_id)
        END,
        COALESCE(
            NULLIF(BTRIM(v_selected->>'discount_profile'), ''),
            a.discount_profile,
            b.discount_profile
        ),
        CASE
            WHEN v_selected ? 'terms_days'
                AND NULLIF(BTRIM(v_selected->>'terms_days'), '') IS NOT NULL
            THEN (v_selected->>'terms_days')::integer
            ELSE COALESCE(a.terms_days, b.terms_days)
        END,
        CASE
            WHEN v_selected ? 'credit_limit'
                AND NULLIF(BTRIM(v_selected->>'credit_limit'), '') IS NOT NULL
            THEN (v_selected->>'credit_limit')::numeric(10,2)
            ELSE COALESCE(a.credit_limit, b.credit_limit)
        END,
        CASE
            WHEN v_selected ? 'stop_credit'
                AND jsonb_typeof(v_selected->'stop_credit') <> 'null'
            THEN (v_selected->>'stop_credit')::boolean
            ELSE COALESCE(a.stop_credit, b.stop_credit, false)
        END,
        CASE
            WHEN v_selected ? 'is_local'
                AND jsonb_typeof(v_selected->'is_local') <> 'null'
            THEN (v_selected->>'is_local')::boolean
            ELSE COALESCE(a.is_local, b.is_local, false)
        END,
        COALESCE(NULLIF(BTRIM(v_selected->>'abn'), ''), a.abn, b.abn),
        CASE
            WHEN v_selected ? 'newsletter_opt_in'
                AND jsonb_typeof(v_selected->'newsletter_opt_in') <> 'null'
            THEN (v_selected->>'newsletter_opt_in')::boolean
            ELSE COALESCE(a.newsletter_opt_in, b.newsletter_opt_in, false)
        END,
        COALESCE(NULLIF(BTRIM(v_selected->>'private_comment'), ''), a.private_comment, b.private_comment),
        COALESCE(
            NULLIF(BTRIM(v_selected->>'statement_comment'), ''),
            a.statement_comment,
            b.statement_comment
        ),
        true,
        COALESCE(NULLIF(BTRIM(p_merged_by), ''), 'customer_merge'),
        COALESCE(
            NULLIF(BTRIM(v_selected->>'musipos_account_code'), ''),
            a.musipos_account_code,
            b.musipos_account_code
        ),
        COALESCE(
            NULLIF(BTRIM(v_selected->>'musipos_barcode_ref'), ''),
            a.musipos_barcode_ref,
            b.musipos_barcode_ref
        )
    )
    RETURNING id, customer_id
    INTO v_new_customer_uuid, v_new_customer_number;

    UPDATE customers
    SET customer_barcode = LPAD(v_new_customer_number::text, 8, '0')
    WHERE id = v_new_customer_uuid;

    SELECT
        COALESCE(
            NULLIF(BTRIM(CONCAT_WS(' ', first_name, surname)), ''),
            NULLIF(BTRIM(business), ''),
            'Merged Customer'
        )
    INTO v_new_customer_name
    FROM customers
    WHERE id = v_new_customer_uuid;

    UPDATE transactions
    SET customer_id = v_new_customer_uuid
    WHERE customer_id IN (p_customer_a, p_customer_b);
    GET DIAGNOSTICS v_tx_count = ROW_COUNT;

    UPDATE transactions
    SET
        customer_id = v_new_customer_uuid,
        customer_name = v_new_customer_name,
        park_name = v_new_customer_name,
        cart_snapshot = jsonb_set(
            jsonb_set(
                COALESCE(cart_snapshot, '{}'::jsonb),
                '{customer_id}',
                to_jsonb(v_new_customer_uuid::text),
                true
            ),
            '{customer_name}',
            to_jsonb(v_new_customer_name),
            true
        )
    WHERE sale_status = 'parked'
      AND (
          customer_id = v_new_customer_uuid
          OR COALESCE(cart_snapshot->>'customer_id', '') IN (p_customer_a::text, p_customer_b::text)
      );
    GET DIAGNOSTICS v_parked_count = ROW_COUNT;

    SELECT to_jsonb(c)
    INTO v_merged_customer_snapshot
    FROM customers AS c
    WHERE c.id = v_new_customer_uuid;

    INSERT INTO customer_merge_audit (
        merged_customer_id,
        source_customer_a_id,
        source_customer_b_id,
        source_customer_a_snapshot,
        source_customer_b_snapshot,
        selected_values,
        merged_customer_snapshot,
        moved_transaction_count,
        moved_parked_count,
        merged_by
    )
    VALUES (
        v_new_customer_uuid,
        p_customer_a,
        p_customer_b,
        to_jsonb(a),
        to_jsonb(b),
        v_selected,
        v_merged_customer_snapshot,
        v_tx_count,
        v_parked_count,
        NULLIF(BTRIM(p_merged_by), '')
    )
    RETURNING id INTO v_audit_id;

    DELETE FROM customers
    WHERE id IN (p_customer_a, p_customer_b);

    RETURN QUERY
    SELECT
        v_audit_id,
        v_new_customer_uuid,
        v_new_customer_number,
        v_tx_count,
        v_parked_count;
END;
$$;
