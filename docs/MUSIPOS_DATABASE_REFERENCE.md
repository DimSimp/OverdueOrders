# MUSIPOS Database & Business Logic Reference

This document provides a complete reference for the MUSIPOS SQL Server database — table schemas, relationships, business workflows, numbering conventions, and critical gotchas. It is designed to give a fresh AI session enough context to safely read from and write to the database.

---

## 1. System Overview

| Property | Value |
|---|---|
| **DBMS** | SQL Server 2008 |
| **Server** | `SERVER\MUSIPOSSQLSRV08` |
| **Database** | `musipos` |
| **User** | `sa` |
| **Driver** | SQL Server Native Client 10.0 |
| **Primary App** | PowerBuilder desktop POS |
| **Triggers** | None on any table — safe for direct writes |
| **Warehouses** | Single: `00001` (char 5) |

The MUSIPOS system is a music retail POS. The PowerBuilder desktop app is the primary consumer of this database. The bulk operations tool (`musipos_bulk`) writes directly to the same tables using the same conventions PowerBuilder expects.

---

## 2. Table Schemas

### 2.1 `ap4itm` — Item Master

Stores every product in the catalogue. Each item has a primary record with `itm_lno = '0000'`.

**Primary Key:** `itm_iid` (char 10) + `itm_lno` (char 4)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `itm_iid` | char(10) | NO | Internal item ID (e.g. `XTAPB1152`) |
| `itm_lno` | char(4) | NO | Line number — always `'0000'` for the main item record |
| `itm_title` | varchar(255) | YES | Full item title |
| `itm_composer` | varchar(150) | YES | Composer / artist name |
| `itm_publisher` | varchar(150) | YES | Publisher / brand |
| `itm_alt` | varchar(20) | YES | Alternate code |
| `itm_sub_dept` | varchar(5) | YES | Sub-department code |
| `itm_status` | varchar(3) | YES | Item status code |
| `itm_titles_in_alb` | varchar(1) | YES | Titles-in-album flag |
| `itm_new_retail_price` | decimal(11,2) | YES | Current retail price (inc. GST) |
| `itm_color` | varchar(1) | YES | Colour code |
| `itm_supplier_id` | varchar(7) | YES | Primary supplier ID → `ap4rsp.rsp_supplier_id` |
| `itm_supplier_iid` | varchar(50) | YES | Supplier's own SKU for this item |
| `itm_input_date` | datetime | YES | Date item was created |
| `itm_last_act_date` | datetime | YES | Last activity date |
| `itm_ameb_title` | varchar(1) | YES | AMEB title flag |
| `itm_barcode_flag` | varchar(1) | YES | Has barcode flag |
| `itm_barcode` | varchar(13) | YES | EAN-13 barcode |
| `itm_brief` | varchar(25) | YES | Brief description (used on invoice lines) |
| `itm_instrument` | varchar(10) | YES | Instrument type |
| `itm_grade` | varchar(10) | YES | Grade level |
| `itm_list` | varchar(10) | YES | List identifier |
| `itm_pic_flag` | varchar(1) | YES | Picture flag |
| `itm_prod_type` | varchar(1) | YES | Product type |
| `itm_trade_id` | varchar(3) | YES | Trade identifier |

**Key relationships:**
- `itm_supplier_id` → `ap4rsp.rsp_supplier_id`
- `itm_iid` → `sp4qpc.qpc_iid` (inventory)
- `itm_iid` → `sp4pop.pop_iid` (PO lines)
- `itm_iid` → `invoice_line.item_id` (sales)

---

### 2.2 `sp4qpc` — Inventory Per Warehouse

One row per item per warehouse. Tracks stock levels, costs, and pricing.

**Primary Key:** `qpc_iid` (char 10) + `qpc_warehouse_id` (char 5)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `qpc_iid` | char(10) | NO | Item ID → `ap4itm.itm_iid` |
| `qpc_warehouse_id` | char(5) | NO | Warehouse ID (always `'00001'`) |
| `qpc_qty_on_hand` | int | YES | **Current stock level** |
| `qpc_qor` | int | YES | Quantity on order (sum of all open PO lines) |
| `qpc_shop_retail` | decimal(11,2) | YES | Shop retail price |
| `qpc_price_increase` | decimal(11,2) | YES | Price increase amount |
| `qpc_min_order` | int | YES | Min reorder qty (-999 = not set) |
| `qpc_max_order` | int | YES | Max reorder qty (-999 = not set) |
| `qpc_location` | varchar(5) | YES | Storage location (default `'001'`) |
| `qpc_location_sub` | varchar(5) | YES | Sub-location (default `'00101'`) |
| `qpc_average_cost` | decimal(11,2) | YES | Average cost (inc. GST) |
| `qpc_last_purchase_cost` | decimal(11,2) | YES | Last purchase cost (inc. GST) — **NOT `qpc_last_cost`** |
| `qpc_warranty` | char(1) | YES | Warranty flag (`'N'`) |
| `qpc_warranty_period` | varchar(2) | YES | Warranty period |
| `qpc_tax_code` | varchar(2) | YES | Tax code (default `'G'` for GST) |
| `qpc_tax_amount` | decimal(11,2) | YES | Tax amount |
| `qpc_alt_supplier_id` | varchar(7) | YES | Alternate supplier ID |
| `qpc_alt_supplier_iid` | varchar(50) | YES | Alternate supplier's SKU |
| `qpc_min_sell` | decimal(11,2) | YES | Minimum sell price |
| `qpc_counted` | int | YES | Stock count flag |
| `qpc_last_invoice_no` | varchar(20) | YES | Last invoice number (updated on receive) |
| `qpc_barcode` | varchar(18) | YES | Barcode (wider than ap4itm) |
| `qpc_category` | varchar(6) | YES | Category code |
| `qpc_last_mod_date` | datetime | YES | Last modification date |
| `qpc_modify_title` | char(1) | YES | Modify title flag |
| `qpc_update_pub_brand` | char(1) | YES | Update publisher/brand flag |
| `qpc_update_inst` | char(1) | YES | Update instrument flag |
| `qpc_upload_to_web` | char(1) | YES | Upload to web flag |
| `qpc_update_title` | char(1) | YES | Update title flag |
| `qpc_update_composer_series` | char(1) | YES | Update composer/series flag |
| `qpc_web_title` | varchar(255) | YES | Web title |
| `qpc_web_specs` | text | YES | Web specifications |
| `qpc_old_web_price` | decimal(11,2) | YES | Old web price |
| `qpc_new_web_price` | decimal(11,2) | YES | New web price |
| `qpc_creation_date` | datetime | YES | Creation date |
| `qpc_studio_rates_item` | char(1) | YES | Studio rates item flag |
| `qpc_static_qoh` | char(1) | YES | Static qty-on-hand flag |
| `qpc_lpc_ex_freight` | decimal(11,4) | YES | Last purchase cost excl. freight (4dp) |
| `qpc_update_status` | char(1) | YES | Update status flag |
| `qpc_update_color` | char(1) | YES | Update color flag |
| `qpc_shopify_pid` | varchar(255) | YES | Shopify product ID |
| `qpc_shopify_variant_id` | varchar(255) | YES | Shopify variant ID |
| `qpc_no_of_updates` | int | YES | Number of updates |
| `qpc_web_cat_id` | varchar(5) | YES | Web category ID |
| `qpc_web_sku` | varchar(50) | YES | Web SKU |
| `qpc_web_disc_reason` | varchar(50) | YES | Web discount reason |
| `qpc_shopify_inv_item_id` | varchar(255) | YES | Shopify inventory item ID |

**How inventory moves:**
- **Sale:** `qpc_qty_on_hand -= qty_purchased`
- **Supplier receive:** `qpc_qty_on_hand += qty_received`
- **PO creation:** `qpc_qor += order_qty`
- **PO line deletion:** `qpc_qor` recalculated from remaining open PO lines

---

### 2.3 `sp4cum` — Customers

Customer master table. Contains contact info, aged balances, and feature flags.

**Primary Key:** `cum_cid` (char 8)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `cum_cid` | char(8) | NO | Customer ID (e.g. `'CASH0001'`, `'JONES001'`) |
| `cum_surname` | nvarchar(20) | YES | Surname |
| `cum_firstname` | nvarchar(20) | YES | First name |
| `cum_address1` | varchar(40) | YES | Address line 1 |
| `cum_address2` | varchar(40) | YES | Address line 2 |
| `cum_city` | varchar(40) | YES | City |
| `cum_state` | varchar(20) | YES | State |
| `cum_pcode` | varchar(10) | YES | Postal code |
| `cum_phone` | nvarchar(15) | YES | Phone |
| `cum_fax` | nvarchar(15) | YES | Fax |
| `cum_mobile` | nvarchar(15) | YES | Mobile |
| `cum_email` | varchar(150) | YES | Email |
| `cum_comment` | nvarchar(60) | YES | Comments |
| `cum_title` | nvarchar(3) | YES | Title (Mr/Ms/etc.) |
| `cum_company_name` | varchar(40) | YES | Company name |
| `cum_discount` | varchar(2) | YES | Discount percentage |
| `cum_terms` | varchar(2) | YES | Payment terms |
| `cum_current_balance` | decimal(11,2) | YES | Current balance |
| `cum_30_balance` | decimal(11,2) | YES | 30-day aged balance |
| `cum_60_balance` | decimal(11,2) | YES | 60-day aged balance |
| `cum_90_balance` | decimal(11,2) | YES | 90-day aged balance |
| `cum_120_balance` | decimal(11,2) | YES | 120+ day aged balance |
| `cum_credit_limit` | decimal(11,2) | YES | Credit limit |
| `cum_invoice` | char(1) | YES | **Y/N flag — NOT an invoice counter** |
| `cum_stop_credit` | varchar(1) | YES | Stop credit flag |
| `cum_sales_order` | char(1) | YES | Has sales order flag |
| `cum_sales_order_printed` | char(1) | YES | Sales order printed flag |
| `cum_sales_order_bo` | char(1) | YES | Sales order backorder flag |
| `cum_abn_no` | varchar(15) | YES | ABN (Australian Business Number) |
| `cum_barcode` | varchar(14) | YES | Customer barcode |
| `cum_date_created` | datetime | YES | Date created |
| `cum_date_modified` | datetime | YES | Date last modified |
| `cum_loyalty_card_no` | varchar(100) | YES | Loyalty card number |
| `cum_musiposwebpassword` | varchar(50) | YES | Web portal password |

*Additional columns exist for shipping address, card details, bank details, studio/teaching fields, custom fields, and communication preferences (80+ columns total). Above are the most commonly referenced.*

**Special customer IDs:**
- `'CASH0001'` — Default cash sale customer
- `'STOCK001'` — Internal stock order customer (used on PO lines for general stock)

---

### 2.4 `ap4rsp` — Suppliers

Supplier master table.

**Primary Key:** `rsp_supplier_id` (char 7)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `rsp_supplier_id` | char(7) | NO | Supplier ID (e.g. `'AMS    '`, `'DADDARI'`) |
| `rsp_name` | varchar(30) | YES | Supplier name |
| `rsp_address_1` | varchar(30) | YES | Address line 1 |
| `rsp_address_2` | varchar(30) | YES | Address line 2 |
| `rsp_city` | varchar(30) | YES | City |
| `rsp_country` | varchar(3) | YES | Country code |
| `rsp_post_code` | varchar(10) | YES | Postal code |
| `rsp_tel` | varchar(15) | YES | Telephone |
| `rsp_fax` | varchar(15) | YES | Fax |
| `rsp_email` | varchar(255) | YES | Email address |
| `rsp_curr_po_no` | int | YES | **Last used PO number** (allocate next = +1) |
| `rsp_acc_no` | varchar(20) | YES | Account number with supplier |
| `rsp_terms` | varchar(3) | YES | Payment terms |
| `rsp_current_balance` | decimal(11,2) | YES | Current balance |
| `rsp_30_balance` | decimal(11,2) | YES | 30-day aged balance |
| `rsp_60_balance` | decimal(11,2) | YES | 60-day aged balance |
| `rsp_90_balance` | decimal(11,2) | YES | 90-day aged balance |
| `rsp_120_balance` | decimal(11,2) | YES | 120+ day aged balance |
| `rsp_retail_factor` | decimal(4,2) | YES | Retail factor / markup |
| `rsp_disc_1` through `rsp_disc_6` | decimal(4,2) | YES | Discount tiers |
| `rsp_settlement_discount` | decimal(4,2) | YES | Settlement discount % |
| `rsp_rebate` | decimal(4,2) | YES | Rebate % |
| `rsp_min_order_value` | int | YES | Minimum order value |
| `rsp_expected_days` | int | YES | Expected delivery days |
| `rsp_po_del_method` | char(1) | YES | PO delivery method (`'F'` = FTP, `'E'` = email) |
| `rsp_inv_inc_tax` | char(1) | YES | Invoices include tax flag |
| `rsp_active` | char(1) | YES | Active flag |
| `rsp_consignment_sup` | char(1) | YES | Consignment supplier flag |
| `rsp_auto_quick_rcv_po` | char(1) | YES | Auto quick-receive PO flag |

*Additional columns exist for distribution address, FTP credentials, bank details, rep info, e-commerce integration (65+ columns total).*

---

### 2.5 `sp4phd` — Purchase Order Headers

One row per PO. No warehouse column — POs are warehouse-agnostic.

**Primary Key:** `phd_supplier_id` (char 7) + `phd_po_no` (int)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `phd_supplier_id` | char(7) | NO | Supplier ID → `ap4rsp.rsp_supplier_id` |
| `phd_po_no` | int | NO | PO number |
| `phd_po_date` | datetime | YES | PO creation date |
| `phd_po_status` | varchar(10) | YES | **Status — see lifecycle below** |
| `phd_ship_info` | varchar(30) | YES | Shipping info |
| `phd_order_info` | varchar(30) | YES | Order info |
| `phd_rcv_date` | datetime | YES | Date received (set when status → RECEIVED) |
| `phd_memo` | text | YES | Memo / notes |
| `phd_po_dest` | char(1) | YES | PO destination (`'S'` = store) |
| `phd_po_prefix` | varchar(20) | YES | PO prefix |

**PO Status Lifecycle:**

```
CURRENT  →  PRINTED  →  PARTRECVD  →  RECEIVED
                  ↓                        ↑
                  └──→ BACKORDER ──────────┘
```

- **CURRENT** — PO created, not yet sent to supplier
- **PRINTED** — Sent to supplier (via email/FTP)
- **PARTRECVD** — Some lines received, others still pending
- **RECEIVED** — All lines fully received (all `sp4pop` lines deleted)
- **BACKORDER** — Items have customer backorders remaining on PO

---

### 2.6 `sp4pop` — Purchase Order Line Items

One row per item per PO. **Critical: received lines are DELETED from this table**, not flagged.

**Primary Key:** `pop_iid` (char 10) + `pop_supplier_id` (char 7) + `pop_po_no` (int) + `pop_cid` (char 8) + `pop_lno` (int)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `pop_iid` | char(10) | NO | Item ID → `ap4itm.itm_iid` |
| `pop_supplier_id` | char(7) | NO | Supplier ID → `ap4rsp.rsp_supplier_id` |
| `pop_po_no` | int | NO | PO number → `sp4phd.phd_po_no` |
| `pop_cid` | char(8) | NO | Customer ID (`'STOCK001'` for stock, real CID for backorders) |
| `pop_lno` | int | NO | Line number (1, 2, 3...) |
| `pop_po_date` | datetime | YES | PO date |
| `pop_supplier_iid` | varchar(50) | YES | Supplier's SKU for this item |
| `pop_qor` | int | YES | Quantity ordered |
| `pop_qty_rcv` | int | YES | Quantity received so far — **NOT `pop_qrc`** |
| `pop_curr_cost` | decimal(11,4) | YES | Current unit cost (inc. GST, 4dp) |
| `pop_curr_price` | decimal(11,2) | YES | Current retail price |
| `pop_supplier_inv_no` | char(10) | YES | Supplier invoice number |
| `pop_freight_charge` | decimal(11,4) | YES | Freight charge per unit |
| `pop_inv_rec_date` | datetime | YES | Invoice received date |
| `pop_curr_order_flag` | char(1) | YES | Current order flag (`'N'` = new, `'Y'` = received) |
| `pop_comment` | char(60) | YES | Comments |
| `pop_qbo` | int | YES | **Quantity on customer backorder** |
| `pop_total_cost` | decimal(11,2) | YES | Total line cost |
| `pop_itm_status` | char(1) | YES | Item status |
| `pop_itm_title` | varchar(255) | YES | Item title (denormalized) |
| `pop_serial_no` | varchar(1) | YES | Serial number flag |
| `pop_user_id` | varchar(4) | YES | User ID |
| `pop_process_status` | char(1) | YES | Process status |
| `pop_prev_po_no` | int | YES | Previous PO number |
| `pop_ext_comment` | varchar(60) | YES | Extended comment |
| `pop_tax_amt` | decimal(11,2) | YES | Tax amount |
| `pop_order_no` | varchar(10) | YES | Customer's sales order number |
| `pop_order_line_no` | int | YES | Customer's order line number |
| `pop_eo` | char(1) | YES | End order flag |
| `pop_computer_id` | varchar(2) | YES | Computer/terminal ID |

**Receiving behaviour:**
- When `qty_received >= qty_ordered` AND no backorder → **DELETE** the row
- When `qty_received >= qty_ordered` AND has backorder → **UPDATE** with adjusted `pop_qor`
- When partially received → **UPDATE** `pop_qty_rcv`

**Customer backorders:**
- `pop_cid` = real customer ID (not `STOCK001`/`CASH0001`)
- `pop_qbo` > 0 = quantity the customer is waiting for
- `pop_order_no` = customer's sales order reference

---

### 2.7 `invoice` — Sales Invoice Headers

One row per sales transaction.

**Primary Key:** `invoice_no` (char 10) + `invoice_date` (datetime)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `invoice_no` | char(10) | NO | Invoice number (allocated from `control_table`) |
| `invoice_date` | datetime | NO | Invoice date (**time = 01:01:01**) |
| `invoice_prefix` | varchar(4) | YES | Prefix (usually empty) |
| `customer_id` | varchar(8) | YES | Customer ID → `sp4cum.cum_cid` |
| `invoice_amount` | decimal(11,2) | YES | Total invoice amount (inc. GST) |
| `collected_amount` | decimal(11,2) | YES | Amount collected |
| `rounding` | decimal(11,2) | YES | Rounding adjustment |
| `change_given` | decimal(11,2) | YES | Change given |
| `ship_name` | varchar(40) | YES | Ship-to name |
| `ship_add1` | varchar(40) | YES | Ship-to address 1 |
| `ship_add2` | varchar(40) | YES | Ship-to address 2 |
| `ship_city` | varchar(40) | YES | Ship-to city |
| `ship_state` | varchar(10) | YES | Ship-to state |
| `ship_pcode` | varchar(10) | YES | Ship-to postcode |
| `ship_via` | char(4) | YES | Shipping method |
| `invoice_comment` | varchar(250) | YES | Comments |
| `user_id` | varchar(4) | YES | User who created (e.g. `'MAN'`) |
| `tax_exempt_no` | varchar(20) | YES | Tax exempt number |
| `freight_charge` | decimal(11,2) | YES | Freight charge |
| `discount_percent` | decimal(11,2) | YES | Discount % |
| `discount_amount` | decimal(11,2) | YES | Discount amount |
| `account_amount` | decimal(11,2) | YES | Account payment amount |
| `cash_amount` | decimal(11,2) | YES | Cash payment |
| `card_type1` | varchar(3) | YES | Card type 1 code (e.g. `'004'`) |
| `card_type1_amount` | decimal(11,2) | YES | Card type 1 amount |
| `card_type2` | varchar(3) | YES | Card type 2 code |
| `card_type2_amount` | decimal(11,2) | YES | Card type 2 amount |
| `card_type3` | varchar(3) | YES | Card type 3 code |
| `card_type3_amount` | decimal(11,2) | YES | Card type 3 amount |
| `cheque_amount` | decimal(11,2) | YES | Cheque amount |
| `quote_no` | char(10) | YES | Related quote number |
| `computer_id` | varchar(2) | YES | Computer/terminal ID (default `'19'`) |
| `invoice_datetime` | datetime | YES | Full datetime with seconds |
| `ship_contact_name` | varchar(50) | YES | Shipping contact |
| `ship_country` | varchar(100) | YES | Shipping country |

*Additional columns exist for interest, email, student attendance, etc.*

---

### 2.8 `invoice_line` — Sales Invoice Line Items

**Primary Key:** `invoice_no` (char 10) + `invoice_date` (datetime) + `invoice_ln_no` (int)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `invoice_no` | char(10) | NO | Invoice number |
| `invoice_date` | datetime | NO | Invoice date |
| `invoice_ln_no` | int | NO | Line number (1, 2, 3...) — **int, NOT varchar** |
| `item_id` | varchar(10) | YES | Item ID → `ap4itm.itm_iid` |
| `qty_purchased` | int | YES | Quantity sold |
| `unit_price` | decimal(11,2) | YES | Unit price (inc. GST) |
| `tax_code` | varchar(2) | YES | Tax code (`'G'` for GST) |
| `tax_rate` | decimal(11,2) | YES | Tax rate (`10.00`) |
| `taxable_price` | decimal(11,2) | YES | Price excl. GST (`total_amount / 11`) |
| `total_amount` | decimal(11,2) | YES | Line total (qty × price) |
| `item_cost` | decimal(11,2) | YES | Item cost (for margin calc) |
| `item_margin` | decimal(11,2) | YES | Margin (price − cost) |
| `supplier_id` | varchar(7) | YES | Supplier ID |
| `supplier_iid` | varchar(50) | YES | Supplier's SKU |
| `disc_percent` | decimal(11,2) | YES | Discount % |
| `disc_amount` | decimal(11,2) | YES | Discount amount |
| `item_description` | varchar(255) | YES | Item description |
| `item_serial_no` | varchar(14) | YES | Serial number |
| `stock_transfer_avg_cost` | decimal(11,2) | YES | Average cost for transfers |

---

### 2.9 `accreceivable` — Accounts Receivable Headers

One AR entry per sales invoice. Links the invoice to its payment record.

**Primary Key:** `accreceivable_no` (char 10) + `accreceivable_date` (datetime)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `accreceivable_no` | char(10) | NO | AR entry number (from `control_table`) |
| `accreceivable_date` | datetime | NO | AR date (= invoice date) |
| `customer_id` | char(8) | YES | Customer ID |
| `invoice_no` | char(10) | YES | Related invoice number |
| `invoice_date` | datetime | YES | Invoice date |
| `trans_type` | char(1) | YES | `'I'` = Invoice |
| `purchase_order_no` | int | YES | PO number (0 for sales) |
| `freight_cost` | decimal(11,2) | YES | Freight cost |
| `invoice_due_date` | datetime | YES | Due date (midnight 00:00:00) |
| `invoice_paid` | char(1) | YES | `'Y'` = paid, `'N'` = on account |

---

### 2.10 `accreceivable_line` — Accounts Receivable Lines

Typically 2 rows per AR entry: line 1 = debit (charge), line 2 = credit (payment).

**Primary Key:** `accreceivable_no` (char 10) + `accreceivable_date` (datetime) + `accreceivable_line_no` (int)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `accreceivable_no` | char(10) | NO | AR number |
| `accreceivable_date` | datetime | NO | AR date |
| `accreceivable_line_no` | int | NO | Line 1 = debit, Line 2 = credit |
| `line_description` | varchar(50) | YES | e.g. `'INVOICE NO: 480585'` or `'CASH'` |
| `debit` | decimal(11,2) | YES | Debit amount (line 1 = total, line 2 = 0) |
| `credit` | decimal(11,2) | YES | Credit amount (line 1 = 0, line 2 = total) |
| `drawer` | varchar(30) | YES | Drawer name |
| `bank` | varchar(30) | YES | Bank code |
| `branch` | varchar(30) | YES | Branch name |
| `cheque_no` | varchar(20) | YES | Cheque number |
| `user_id` | varchar(4) | YES | User ID |
| `user_date` | datetime | YES | Entry date |
| `discount_pct` | decimal(4,2) | YES | Discount % |
| `discount_amount` | decimal(11,2) | YES | Discount amount |
| `balance` | decimal(11,2) | YES | Running balance (line 1 = total, line 2 = 0) |
| `payment_type` | varchar(1) | YES | `'P'` = payment |
| `cash_amount` | decimal(11,2) | YES | Cash portion |
| `card_type1` | varchar(3) | YES | Card type 1 |
| `card_type1_amount` | decimal(11,2) | YES | Card type 1 amount |
| `card_type2` | varchar(3) | YES | Card type 2 |
| `card_type2_amount` | decimal(11,2) | YES | Card type 2 amount |
| `card_type3` | varchar(3) | YES | Card type 3 |
| `card_type3_amount` | decimal(11,2) | YES | Card type 3 amount |
| `cheque_amount` | decimal(11,2) | YES | Cheque amount |
| `layby_credit` | decimal(11,2) | YES | Lay-by credit |
| `deposit_credit` | decimal(11,2) | YES | Deposit credit |
| `voucher_credit` | decimal(11,2) | YES | Voucher credit |
| `trans_posted` | char(1) | YES | Posted flag (`'N'`) |
| `computer_id` | varchar(2) | YES | Terminal ID |
| `ar_journal_no` | varchar(10) | YES | AR journal number |
| `layby_no` | varchar(10) | YES | Lay-by number |

---

### 2.11 `accpayable` — Accounts Payable Headers

One AP entry per supplier invoice received. **Has NO `invoice_amount` column** — the amount is on `accpayable_line`.

**Primary Key:** `accpayable_no` (char 10) + `accpayable_date` (datetime)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `accpayable_no` | char(10) | NO | AP entry number (from `control_table`) |
| `accpayable_date` | datetime | NO | AP date (01:01:01 time) |
| `supplier_id` | char(7) | YES | Supplier ID |
| `invoice_no` | varchar(20) | YES | Supplier's invoice number |
| `invoice_date` | datetime | YES | Supplier's invoice date |
| `trans_type` | char(1) | YES | `'I'` = Invoice, `'P'` = Payment |
| `purchase_order_no` | int | YES | PO number |
| `freight_cost` | decimal(11,2) | YES | Freight cost |
| `invoice_due_date` | datetime | YES | Due date (end of next month, 01:01:01) |
| `invoice_paid` | char(1) | YES | `'N'` = unpaid |
| `rebate` | decimal(11,2) | YES | Rebate amount |
| `exchange_rate` | decimal(11,8) | YES | Exchange rate (default `1.0`) |

---

### 2.12 `accpayable_line` — Accounts Payable Lines

Typically 1 row per AP entry — a credit for the full invoice amount.

**Primary Key:** `accpayable_no` (char 10) + `accpayable_date` (datetime) + `accpayable_line_no` (int)

| Column | Type | Nullable | Purpose |
|---|---|---|---|
| `accpayable_no` | char(10) | NO | AP number |
| `accpayable_date` | datetime | NO | AP date |
| `accpayable_line_no` | int | NO | Line number (typically 1) |
| `line_description` | varchar(50) | YES | e.g. `'INVOICE NO: INV12345'` |
| `debit` | decimal(11,2) | YES | Debit (0 for invoice entry) |
| `credit` | decimal(11,2) | YES | **Credit = invoice total (inc. GST)** |
| `drawer` | varchar(30) | YES | Drawer |
| `bank` | varchar(30) | YES | Bank |
| `branch` | varchar(30) | YES | Branch |
| `cheque_no` | varchar(20) | YES | Cheque number |
| `user_id` | varchar(4) | YES | User ID |
| `user_date` | datetime | YES | Entry date |
| `discount_pct` | decimal(4,2) | YES | Discount % |
| `discount_amount` | decimal(11,2) | YES | Discount amount |
| `balance` | decimal(11,2) | YES | Balance (= credit amount) |
| `rebate_amt` | decimal(11,2) | YES | Rebate amount |
| `credit_req_amt` | decimal(11,2) | YES | Credit request amount |

---

### 2.13 `accpayable_po` — AP Purchase Order Detail

Audit trail of received items. One row per item received on a supplier invoice. Populates the "Supplier Invoice History Transactions" view in PowerBuilder.

**Primary Key:** `accpayable_no` + `accpayable_date` + `accpayable_line_no`

| Column | Type | Purpose |
|---|---|---|
| `accpayable_no` | char(10) | AP entry number |
| `accpayable_date` | datetime | AP date |
| `accpayable_line_no` | int | Sequential line (1, 2, 3... per AP) |
| `pop_iid` | char(10) | Item ID |
| `pop_supplier_id` | char(7) | Supplier ID |
| `pop_po_no` | int | PO number |
| `pop_cid` | char(8) | Customer ID from PO line |
| `pop_lno` | int | PO line number |
| `pop_po_date` | datetime | PO date |
| `pop_supplier_iid` | varchar(50) | Supplier's SKU |
| `pop_qor` | int | Quantity ordered |
| `pop_qty_rcv` | int | Quantity received (this invoice) |
| `pop_curr_cost` | decimal(11,4) | Unit cost (inc. GST) |
| `pop_curr_price` | decimal(11,2) | Retail price |
| `pop_supplier_inv_no` | char(10) | Supplier invoice number |
| `pop_freight_charge` | decimal(11,4) | Freight per unit |
| `pop_inv_rec_date` | datetime | Date received |
| `pop_curr_order_flag` | char(1) | `'Y'` = received |
| `pop_comment` | char(60) | Comment |
| `pop_qbo` | int | Backorder quantity |
| `pop_total_cost` | decimal(11,2) | Total line cost |
| `pop_itm_status` | char(1) | Item status |
| `pop_itm_title` | varchar(255) | Item title |
| `pop_serial_no` | varchar(1) | Serial number flag |
| `pop_user_id` | varchar(4) | User ID |
| `pop_process_status` | char(1) | Process status |
| `pop_tax_amt` | decimal(11,2) | GST amount |

---

### 2.14 `control_table` — Number Allocation

Stores sequential counters for invoice, AR, and AP numbering.

| Column | Type | Purpose |
|---|---|---|
| `control_id` | char(20) | Counter name |
| `control_info` | varchar(40) | Counter value (stored as string) |

**Known control IDs:**

| control_id | Purpose | Semantics |
|---|---|---|
| `INVOICE_NO` | Sales invoice numbers | **LAST USED** (not next-to-use) |
| `ACCRECEIVABLE_NO` | AR entry numbers | **LAST USED** |
| `ACCPAYABLE_NO` | AP entry numbers | **LAST USED** |

---

### 2.15 `stock_audit` — Inventory Movement Audit

Records every inventory change with before/after snapshots.

| Column | Type | Purpose |
|---|---|---|
| `item_id` | varchar(10) | Item ID |
| `trans_no` | char(10) | Per-item sequential number (zero-padded) |
| `trans_date` | datetime | Transaction date (`GETDATE()`) |
| `user_id` | varchar(4) | User ID |
| `trans_type` | char(1) | `'S'` = sale, `'P'` = purchase/receive |
| `old_qty_on_hand` | int | Qty before |
| `qty_processed` | int | Change (+ve = stock out, −ve = stock in) |
| `new_qty_on_hand` | int | Qty after |
| `old_cost` | decimal(11,2) | Cost before |
| `new_cost` | decimal(11,2) | Cost after |
| `customer_id` | varchar(8) | Customer (sales) or blank (receives) |
| `trans_doc_no` | varchar(10) | Invoice/PO number |
| `trans_doc_date` | datetime | Document date |

---

## 3. Table Relationships

```
┌─────────────────────────────────────────────────────────────┐
│                      SALES FLOW                             │
│                                                             │
│  control_table ──→ invoice_no                               │
│  control_table ──→ accreceivable_no                         │
│                                                             │
│  sp4cum ←── invoice.customer_id                             │
│                  │                                          │
│                  ├──→ invoice_line (1:N)                     │
│                  │        │                                  │
│                  │        ├──→ ap4itm (item lookup)          │
│                  │        └──→ sp4qpc (decrement qty)        │
│                  │                                          │
│                  └──→ accreceivable (1:1)                    │
│                           └──→ accreceivable_line (1:N)      │
│                                (line 1: debit, line 2: credit)│
│                                                             │
│  stock_audit ←── (per inventory change)                     │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    PURCHASING FLOW                           │
│                                                             │
│  ap4rsp.rsp_curr_po_no ──→ PO number allocation             │
│  control_table ──→ accpayable_no                            │
│                                                             │
│  ap4rsp ←── sp4phd.phd_supplier_id (PO header)              │
│                 │                                           │
│                 └──→ sp4pop (PO lines, 1:N)                  │
│                          │                                  │
│                          ├──→ ap4itm (item lookup)           │
│                          ├──→ sp4qpc (increment qty on recv) │
│                          └──→ sp4cum (backorder customer)    │
│                                                             │
│  On receive:                                                │
│    sp4pop ──→ DELETE (fully received)                        │
│    sp4pop ──→ UPDATE (partial/backorder)                     │
│    sp4qpc ──→ UPDATE (qty_on_hand += received)              │
│    sp4phd ──→ UPDATE (status → RECEIVED/BACKORDER)          │
│                                                             │
│  accpayable (1:1 per supplier invoice)                      │
│      ├──→ accpayable_line (1 row: credit for total)          │
│      └──→ accpayable_po (1:N: per-item audit trail)          │
│                                                             │
│  stock_audit ←── (per inventory change)                     │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                   ITEM RELATIONSHIPS                        │
│                                                             │
│  ap4itm.itm_iid ←→ sp4qpc.qpc_iid (inventory)              │
│  ap4itm.itm_iid ←→ sp4pop.pop_iid (PO lines)               │
│  ap4itm.itm_iid ←→ invoice_line.item_id (sales)             │
│  ap4itm.itm_supplier_id ──→ ap4rsp.rsp_supplier_id          │
└─────────────────────────────────────────────────────────────┘
```

---

## 4. Business Workflows

### 4.1 Sales Invoice Creation

Creates a sales invoice, decrements inventory, and writes AR records.

**Table writes in order:**

| Step | Table | Operation | Key Detail |
|---|---|---|---|
| 1 | `control_table` | SELECT+UPDATE | Allocate invoice number (UPDLOCK) |
| 2 | `control_table` | SELECT+UPDATE | Allocate AR number (UPDLOCK) |
| 3 | `invoice` | INSERT | Header with payment split |
| 4 | `invoice_line` | INSERT × N | One row per item |
| 5 | `sp4qpc` | UPDATE × N | `qpc_qty_on_hand -= qty_purchased` |
| 6 | `stock_audit` | INSERT × N | Trans type `'S'`, `qty_processed = +qty` |
| 7 | `accreceivable` | INSERT | Links to invoice, `trans_type = 'I'` |
| 8 | `accreceivable_line` | INSERT (line 1) | DEBIT: charge to customer (balance = total) |
| 9 | `accreceivable_line` | INSERT (line 2) | CREDIT: payment received (balance = 0) |

**Payment type determines:**
- `CASH` → `cash_amount = total`, `invoice_paid = 'Y'`
- `CARD` → `card_type1_amount = total`, `invoice_paid = 'Y'`
- `ACCOUNT` → `account_amount = total`, `invoice_paid = 'N'`

---

### 4.2 Supplier Invoice Receiving

Receives stock against a PO, increments inventory, and creates AP records.

**Table writes in order:**

| Step | Table | Operation | Key Detail |
|---|---|---|---|
| 1 | `ap4rsp` | SELECT+UPDATE | Allocate PO number if auto-creating (UPDLOCK) |
| 2 | `sp4phd` | INSERT | Only if auto-creating PO |
| 3 | `sp4pop` | INSERT × N | Only if auto-creating PO lines |
| 4 | `sp4pop` | SELECT | Read PO line snapshot (before modify/delete) |
| 5 | `sp4pop` | SELECT | Check ALL POs for customer backorders on item |
| 6 | `sp4pop` | DELETE or UPDATE | Per line: delete if fully received, update if partial/backorder |
| 7 | `sp4qpc` | UPDATE × N | `qpc_qty_on_hand += qty_received` |
| 8 | `sp4qpc` | INSERT | Create missing qpc record if item has none |
| 9 | `sp4qpc` | UPDATE | `qpc_last_purchase_cost = cost_inc_gst` |
| 10 | `stock_audit` | INSERT × N | Trans type `'P'`, `qty_processed = −qty` (negative = stock in) |
| 11 | `sp4pop` | SELECT COUNT | Check remaining lines per PO |
| 12 | `sp4phd` | UPDATE | Status → `RECEIVED` or `BACKORDER` |
| 13 | `control_table` | SELECT+UPDATE | Allocate AP number (UPDLOCK) |
| 14 | `accpayable` | INSERT | Header: `trans_type = 'I'`, `invoice_paid = 'N'` |
| 15 | `accpayable_line` | INSERT | 1 row: CREDIT for invoice total |
| 16 | `accpayable_po` | INSERT × N | 1 row per received item (audit trail) |

**Cost handling:**
- CSV/PDF provides unit cost EXCLUSIVE of GST
- System multiplies by 1.1 for INCLUSIVE storage
- `sp4qpc.qpc_last_purchase_cost` = cost inc. GST
- `sp4pop.pop_curr_cost` = cost inc. GST (4 decimal places)

**AP due date:** End of the NEXT month from invoice date (e.g. Feb 15 → Mar 31)

---

### 4.3 Purchase Order Creation

Creates or updates POs for items with stock shortfall.

**Table writes in order:**

| Step | Table | Operation | Key Detail |
|---|---|---|---|
| 1 | `sp4phd` | SELECT | Find existing CURRENT PO for supplier |
| 2a | `sp4pop` | SELECT | Check if item already on PO |
| 2b | `sp4pop` | UPDATE | If yes: `pop_qor += order_qty` |
| 2c | `sp4pop` | INSERT | If no: new line (`pop_cid = 'STOCK001'`) |
| 3 | `sp4qpc` | UPDATE | `qpc_qor += order_qty` |
| — | *Or if no CURRENT PO:* | | |
| 4 | `ap4rsp` | SELECT+UPDATE | Allocate PO number (UPDLOCK) |
| 5 | `sp4phd` | INSERT | New PO header, status = `'CURRENT'` |
| 6 | `sp4pop` | INSERT × N | New PO lines |
| 7 | `sp4qpc` | UPDATE × N | `qpc_qor += order_qty` |

---

### 4.4 PO Sending

Sends POs to suppliers and updates status.

| Step | Table | Operation | Key Detail |
|---|---|---|---|
| 1 | `sp4phd` + `ap4rsp` | SELECT | Get PO header + supplier contact info |
| 2 | `sp4pop` + `ap4itm` | SELECT | Get PO lines + item details |
| 3 | — | Email/FTP | Send PDF or CSV to supplier |
| 4 | `sp4phd` | UPDATE | `phd_po_status = 'PRINTED'` (only if was `'CURRENT'`) |

---

### 4.5 Invoice Rollback

Reverses a sales invoice: restores inventory, deletes AR, deletes invoice.

| Step | Table | Operation | Key Detail |
|---|---|---|---|
| 1 | `invoice` | SELECT | Verify invoice exists |
| 2 | `invoice_line` | SELECT | Get lines for inventory restore |
| 3 | `accreceivable` | SELECT | Find linked AR record |
| 4 | `accreceivable_line` | DELETE | Remove AR lines |
| 5 | `accreceivable` | DELETE | Remove AR header |
| 6 | `sp4qpc` | UPDATE × N | `qpc_qty_on_hand += qty_purchased` (restore) |
| 7 | `invoice_line` | DELETE | Remove invoice lines |
| 8 | `invoice` | DELETE | Remove invoice header |
| 9 | `control_table` | UPDATE | Decrement counters (only if was the last invoice) |

---

## 5. Number Allocation Pattern

All sequential numbering uses a **read-lock-increment** pattern with `UPDLOCK + HOLDLOCK` for atomicity within a transaction.

### control_table (Invoice, AR, AP numbers)

```sql
-- Step 1: Read with exclusive lock (prevents concurrent reads)
SELECT CAST(control_info AS int)
FROM control_table WITH (UPDLOCK, HOLDLOCK)
WHERE control_id = 'INVOICE_NO'

-- Step 2: In application code
next_no = last_used + 1

-- Step 3: Update counter
UPDATE control_table
SET control_info = CAST(@next_no AS varchar)
WHERE control_id = 'INVOICE_NO'

-- Step 4: Use next_no as the new invoice number
```

**Semantics: `control_info` stores LAST USED, not next-to-use.** PowerBuilder reads, adds 1, uses that value, then updates.

### PO Numbers (per-supplier, from ap4rsp)

```sql
-- Read with lock
SELECT rsp_curr_po_no
FROM ap4rsp WITH (UPDLOCK, HOLDLOCK)
WHERE rsp_supplier_id = @supplier_id

-- Compute +1, use it, then update
UPDATE ap4rsp
SET rsp_curr_po_no = @new_po_no
WHERE rsp_supplier_id = @supplier_id
```

---

## 6. SKU Resolution Cascade

When resolving a SKU string to an item, the system tries three columns in order:

| Priority | Column | Table | Purpose |
|---|---|---|---|
| 1st | `itm_iid` | `ap4itm` | Internal MUSIPOS item ID |
| 2nd | `itm_supplier_iid` | `ap4itm` | Supplier's SKU for the item |
| 3rd | `itm_barcode` | `ap4itm` | EAN-13 barcode |

All lookups filter on `itm_lno = '0000'` (main item record only).

An optional `supplier_id` parameter adds `AND itm_supplier_id = ?` to disambiguate items that exist under multiple suppliers.

The query JOINs `ap4itm` with `sp4qpc` to also return inventory data (qty on hand, costs).

**Fuzzy matching fallback:**
1. Exact uppercase comparison
2. OCR confusion normalization (0↔O, 1↔I/l, 5↔S, 8↔B)
3. `difflib.SequenceMatcher` ratio > 0.75 with length difference ≤ 2

---

## 7. Critical Conventions & Gotchas

### Date/Time Formatting
| Context | Time Component | Example |
|---|---|---|
| Invoice/document dates | `01:01:01` | `2026-03-15 01:01:01.000` |
| AR due dates | `00:00:00` (midnight) | `2026-03-15 00:00:00.000` |
| AP due dates | End of NEXT month, `01:01:01` | Feb invoice → `2026-03-31 01:01:01.000` |
| `invoice_datetime` | Actual `NOW()` | `2026-03-15 14:32:05.123` |
| `GETDATE()` fields | Actual current time | |

PowerBuilder hard-codes the `01:01:01` time component. Any direct writes must match this convention or PowerBuilder may not find the records.

### Field Name Traps
| Wrong | Correct | Table |
|---|---|---|
| `pop_qrc` | `pop_qty_rcv` | `sp4pop` |
| `qpc_last_cost` | `qpc_last_purchase_cost` | `sp4qpc` |
| `invoice_ln_no` is varchar | `invoice_ln_no` is **int** | `invoice_line` |
| `accpayable.invoice_amount` | **Column does not exist** | `accpayable` |
| `sp4phd` has warehouse | `sp4phd` has **NO warehouse column** | `sp4phd` |
| `cum_invoice` is a counter | `cum_invoice` is **char(1) Y/N flag** | `sp4cum` |

### ID Formats
| ID | Type | Format | Example |
|---|---|---|---|
| Warehouse ID | char(5) | Zero-padded | `'00001'` (NOT `'01'`) |
| Customer ID | char(8) | Alphanumeric | `'CASH0001'`, `'JONES001'` |
| Supplier ID | char(7) | Alphanumeric | `'AMS    '`, `'DADDARI'` |
| Invoice number | char(10) | Numeric string | `'480585    '` |
| Item ID | char(10) | Alphanumeric | `'XTAPB1152 '` |
| User ID | varchar(4) | Short code | `'MAN'` |
| Computer ID | varchar(2) | Numeric | `'19'` |

### GST (Australian Tax)
- Rate: 10% (multiplier `1.1` for inclusive pricing)
- All prices in the database are **inclusive of GST**
- CSV/PDF supplier invoices provide costs **exclusive** of GST
- System multiplies by 1.1 before storing
- Tax code: `'G'` (GST)
- Taxable price = `total_amount / 11` (back-calculate ex-GST component)

### Empty Strings vs NULL
PowerBuilder expects empty strings `''` rather than `NULL` for most text fields. When inserting records, default text fields to `''` unless there's a real value.

### Transaction Management
- All bulk operations use a transaction context manager (`autocommit=False`)
- `transaction()` — commits on success, rolls back on exception
- `dry_run_transaction()` — always rolls back (preview mode)
- Number allocation uses `UPDLOCK + HOLDLOCK` within the transaction for atomicity

### PO Line Deletion on Receive
When a PO line is fully received (`qty_received >= qty_ordered`) and there's no customer backorder, the row is **DELETED from `sp4pop`** — it is NOT flagged or marked. The `accpayable_po` table preserves the audit trail of what was received.

### Special Customer IDs
- `'CASH0001'` — Default walk-in / cash sale customer
- `'STOCK001'` — Used on PO lines (`pop_cid`) for general stock orders (no customer attached)
- Any other `pop_cid` with `pop_qbo > 0` indicates a customer backorder

---

## 8. Sample Data Reference

| Data Point | Value |
|---|---|
| Recent invoice range | ~480400–480600 |
| Default customer | `CASH0001` |
| Default user | `MAN` |
| Computer ID | `19` |
| Warehouse | `00001` |
| Recent AP range | ~99997–99999 |
| Sample suppliers | AMS, ELECTRI, DADDARI |

---

## 9. Transaction Pattern Templates

### Creating a Sales Invoice

```sql
-- 1. Allocate invoice number
SELECT CAST(control_info AS int) FROM control_table WITH (UPDLOCK, HOLDLOCK)
  WHERE control_id = 'INVOICE_NO'
-- app: next_no = result + 1
UPDATE control_table SET control_info = CAST(@next_no AS varchar)
  WHERE control_id = 'INVOICE_NO'

-- 2. Insert invoice header
INSERT INTO invoice (invoice_no, invoice_date, customer_id, invoice_amount,
  collected_amount, cash_amount, card_type1, card_type1_amount, user_id,
  computer_id, invoice_datetime, ...)
VALUES (@next_no, @date_with_010101, @customer_id, @total, @total,
  @cash_amt, '004', @card_amt, @user_id, '19', GETDATE(), ...)

-- 3. Insert invoice lines (per item)
INSERT INTO invoice_line (invoice_no, invoice_date, invoice_ln_no, item_id,
  qty_purchased, unit_price, total_amount, item_cost, item_margin,
  item_description, supplier_id, supplier_iid,
  tax_code, tax_rate, taxable_price, ...)
VALUES (@next_no, @date, @line_no_int, @itm_iid, @qty, @price,
  @line_total, @cost, @margin, @desc, @supplier_id, @supplier_iid,
  'G', 10.00, @line_total / 11, ...)

-- 4. Decrement inventory
UPDATE sp4qpc SET qpc_qty_on_hand = qpc_qty_on_hand - @qty
  WHERE qpc_iid = @itm_iid AND qpc_warehouse_id = '00001'

-- 5. Allocate AR number (same UPDLOCK pattern)
-- 6. Insert accreceivable header
-- 7. Insert accreceivable_line (2 rows: debit + credit)
```

### Receiving a Supplier Invoice

```sql
-- For each item on the invoice:

-- 1. Check for customer backorders (across ALL POs)
SELECT pop_cid, pop_qbo, pop_order_no, pop_po_no FROM sp4pop
  WHERE pop_iid = @itm_iid
    AND pop_cid IS NOT NULL AND RTRIM(pop_cid) <> ''
    AND pop_qbo > 0

-- 2. Update or delete PO line
-- If fully received and no backorder:
DELETE FROM sp4pop
  WHERE pop_supplier_id = @sid AND pop_po_no = @po AND pop_iid = @iid
-- If partial:
UPDATE sp4pop SET pop_qty_rcv = pop_qty_rcv + @qty_rcv,
  pop_supplier_inv_no = @inv_no, pop_curr_cost = @cost_inc_gst
  WHERE pop_supplier_id = @sid AND pop_po_no = @po AND pop_iid = @iid

-- 3. Increment inventory
UPDATE sp4qpc SET qpc_qty_on_hand = qpc_qty_on_hand + @qty_rcv,
  qpc_last_purchase_cost = @cost_inc_gst,
  qpc_last_invoice_no = @inv_no, qpc_last_mod_date = @date
  WHERE qpc_iid = @iid AND qpc_warehouse_id = '00001'

-- 4. Update PO status (after processing all lines)
-- Count remaining lines:
SELECT COUNT(*) FROM sp4pop
  WHERE pop_supplier_id = @sid AND pop_po_no = @po
-- If 0: UPDATE sp4phd SET phd_po_status = 'RECEIVED', phd_rcv_date = GETDATE()
-- If has backorders: UPDATE sp4phd SET phd_po_status = 'BACKORDER'

-- 5. Create AP record (UPDLOCK pattern for AP number)
INSERT INTO accpayable (...) VALUES (...)
INSERT INTO accpayable_line (...) VALUES (...) -- 1 credit row
INSERT INTO accpayable_po (...) VALUES (...)   -- 1 row per received item
```
