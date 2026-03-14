# Scarlett AIO — Bug Fixes & Planned Improvements

## Bugs

- [ ] **Results tab: items duplicate on fresh session load**
  Items sometimes appear twice in the results table when loading a `.scar` session file.
  They disappear after a manual refresh. Root cause unknown — likely a double-populate
  triggered during the session load / tab activation sequence.

- [ ] **Deletion modal: date sort is wrong order**
  Sorting by Date column should show the most recent bookings first
  (e.g. 14/03/2026 10:00am before 14/03/2026 9:55am).
  Currently the date and time columns are sorted independently; they need to be
  combined into a single datetime key for correct ordering.

- [ ] **Reprint from Manage Shipments modal should actually print**
  Currently "Reprint Label" opens the PDF in the default viewer.
  It should send the label directly to the Brother QL label printer using the
  same `print_label()` path as the original booking, with the same per-courier
  settings (scale, split, etc.).

- [ ] **Label print size differs between build and development**
  Some labels print smaller when running from the packaged exe compared to
  running from source. Likely cause: `label_settings.json` is not being found
  in the correct location when frozen, so default settings are used instead of
  the saved per-courier values. Needs investigation.

---

## Improvements

- [ ] **Settings window — configurable file paths**
  Add a Settings screen (accessible from the header or a menu) where users can
  configure the following paths without editing `config.json` manually:
  - Bookings directory (network share path)
  - Lists / CSV output directory
  - Session / snapshot (`.scar`) directory
  - *(Additional paths to confirm — user selected "Other")*
  Changes should be saved back to `config.json` immediately.

- [ ] **Move `config.json` to AppData (survives updates)**
  Currently `config.json` lives next to the exe and must be manually copied
  whenever a user installs a new version.
  Solution: store it at `%APPDATA%\ScarlettAIO\config.json` so it persists
  across updates automatically.
  Requires updating `src/config.py` path resolution and migrating any existing
  `config.json` found next to the exe on first launch.

- [ ] **Refresh single order after note added via fulfilment screen**
  After saving a note to a Neto order from the order detail / fulfilment screen,
  the row in the Orders or Results tab should refresh to reflect the new note
  without requiring a full reload of all orders.