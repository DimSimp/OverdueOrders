# Future: Cloud-Based eBay Dispatch Retry Queue

## Background

When eBay's Seller API is completely unavailable (both the API and the portal are inaccessible), orders dispatched via Neto cannot be updated on eBay at the time of dispatch. The current solution presents staff with a popup linking directly to the eBay order in Seller Hub so it can be updated manually — this works in the common case where the portal is accessible even when the API is down.

In a full outage scenario, however, manual updates are also blocked. To handle this edge case, a cloud-based retry queue would allow queued orders to be automatically fulfilled on eBay once the API recovers, without requiring any PC to be on at the time of recovery.

## Proposed Architecture

### Queue Storage — Supabase

Supabase is already configured and used in this project. A new table `pending_ebay_syncs` would store orders that failed to sync:

```sql
create table pending_ebay_syncs (
    ebay_order_id   text primary key,
    tracking_number text not null default '',
    carrier         text not null default '',
    added_at        timestamptz not null default now(),
    last_attempted  timestamptz,
    attempt_count   int not null default 0
);
```

- The desktop app inserts a row on sync failure.
- Rows are deleted on successful retry.
- Using `ebay_order_id` as the primary key prevents duplicates naturally.

### Retry Execution — GitHub Actions (Scheduled)

A small private GitHub repository (separate from this one) would contain:

1. A Python script (`retry_ebay_syncs.py`) that:
   - Reads all rows from `pending_ebay_syncs`
   - For each row: checks eBay order status first (avoid double-fulfillment if staff already did it manually)
   - If not yet fulfilled: calls `create_shipping_fulfillment()` with stored tracking data
   - On success: deletes the row from Supabase
   - On failure: updates `last_attempted` and increments `attempt_count`

2. A GitHub Actions workflow (`.github/workflows/retry.yml`) scheduled with a cron trigger (e.g. every 30 minutes):

```yaml
on:
  schedule:
    - cron: '*/30 * * * *'
  workflow_dispatch:  # also allows manual trigger from GitHub UI
```

All credentials (eBay OAuth tokens, Supabase key) are stored as GitHub Secrets — not in code.

### Cost

- **Supabase**: Free tier (this project already uses it)
- **GitHub Actions**: 2,000 minutes/month on private repos. Each retry run takes under one minute. At 30-minute intervals that's ~48 runs/day = ~1,440 runs/month = ~720 minutes/month — well within the free limit.

## Key Design Considerations

### Always check order status before retrying
Call `get_order_status()` before `create_shipping_fulfillment()`. If the order is already `FULFILLED` (staff manually updated it via the popup link), skip the API call and remove it from the queue. This prevents duplicate fulfillments.

### Store tracking data at the time of failure
The queue row must include the tracking number and carrier from the original dispatch, so the retry submits consistent data rather than a blank fulfillment.

### Handle stale entries
eBay's Fulfillment API has a window (typically 30 days) for creating shipping fulfillments. Entries older than this window should be flagged rather than retried silently — they may require manual resolution.

### Process one order at a time, persist immediately
Remove each row from Supabase the moment its retry succeeds. Don't batch-delete at the end — if the script is interrupted mid-run, completed orders won't be retried unnecessarily.

## Why This Wasn't Implemented Yet

Full API + portal outages are rare (a few times per year), and the popup approach handles the common case (API down, portal accessible) which covers ~90% of incidents. The cloud infrastructure adds meaningful complexity and a new external dependency for an edge case that hasn't caused significant operational impact.

Revisit if full outages become more frequent or if the manual fallback becomes impractical (e.g. high order volume during an outage).
