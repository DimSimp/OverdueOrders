# Smartphone Photo Import — Research Notes

## The Goal
Allow a staff member to photograph physical invoices with their phone and have the images appear immediately in the app for AI parsing — without any external service accounts, mobile app installs, or major code restructure. Multiple photos should be capturable in one session (multi-page invoices).

---

## Option A — Local HTTP Server + QR Code ⭐ Recommended

### How it works
1. User clicks a new "Scan Invoice with Phone" button in the invoice tab.
2. The app starts a tiny built-in HTTP server on a random local port (e.g. `http://192.168.1.5:8765`).
3. A QR code is displayed in a small dialog window pointing to that URL.
4. User scans the QR code with their phone camera — it opens a webpage in the phone's browser.
5. The webpage has a multi-file camera input (`<input type="file" accept="image/*" multiple>`). On mobile browsers this opens the camera directly or the photo library.
6. User takes/selects photos, taps Upload.
7. The server receives the images, saves them to a temp folder, and signals the app.
8. The QR dialog closes; images are queued into the normal AI parse pipeline exactly as if the user had selected local files.

### Why this is the best fit
- **Zero phone setup** — works in any mobile browser (Safari/Chrome), no app to install.
- **Local network only** — no internet required, no cloud accounts, no data leaves the premises.
- **Multiple photos in one go** — `multiple` attribute on the file input handles multi-page invoices.
- **Integrates cleanly** — images drop into the existing `_IMAGE_EXTENSIONS` → `_image_to_temp_pdf()` → AI parse path with no new concepts.
- **All dependencies are stdlib or already installed** — `http.server` (stdlib), `threading` (stdlib), `qrcode` (one small package), `Pillow` (already installed).

### What the phone page looks like
```html
<form method="POST" action="/upload" enctype="multipart/form-data">
  <input type="file" name="photos" accept="image/*" multiple capture="environment">
  <button type="submit">Upload Photos</button>
</form>
```
`capture="environment"` hints to open the rear camera immediately on Android. iOS ignores it and shows the standard photo picker instead, but still works.

### Dependencies to add
```
qrcode[pil]>=7.4
```
`Pillow` is already installed. `http.server` is stdlib.

### Limitations
- Phone and PC must be on **the same Wi-Fi network** (or same Ethernet switch). This is true in almost all office/shop environments.
- If Windows Firewall blocks the port, a one-time prompt to allow it will appear.
- The upload page is very plain — functional but not styled. This is intentional (keep it simple/fast).
- Very large photos (12+ MP) will work but may slow AI parsing; the conversion step could optionally downscale to 2000px max before sending to OpenAI.

### Rough implementation scope
- ~100–150 lines of new Python (HTTP server class + QR dialog)
- One new button in the invoice tab controls row
- No changes to the parse pipeline

---

## Option B — Watched Folder via Cloud Sync

### How it works
1. User configures a sync folder in config.json (e.g. `C:\Users\...\OneDrive\InvoicePhotos`).
2. App polls or watches the folder for new image files.
3. User takes photos on their phone → they automatically sync to that folder via OneDrive/Dropbox/Google Drive.
4. App detects new files and prompts "3 new images detected — import now?"

### Pros
- Works from anywhere, not just same Wi-Fi.
- Very reliable (sync handles retries, compression, etc.).
- Zero code for the actual file transfer part.

### Cons
- Requires a cloud account (OneDrive/Dropbox/Google Drive) to be set up and syncing on both devices.
- Sync delay: typically 5–30 seconds after taking the photo.
- Folder watch requires either polling (simple but slightly laggy) or `watchdog` library.
- Not "instant" — the user has to wait for sync before the app can see the images.
- Clutters a cloud sync folder with invoice photos unless it's cleaned up automatically.

### Best for
Situations where the phone and PC are on different networks, or where staff take photos throughout the day and batch-process them later.

---

## Option C — Telegram Bot

### How it works
1. Create a free Telegram bot via @BotFather (takes 2 minutes, gives an API token).
2. App polls `api.telegram.org/getUpdates` for new messages to the bot.
3. Staff send invoice photos to the bot from any Telegram account.
4. App downloads the photos and queues them for parsing.

### Pros
- Works anywhere (internet-based, no Wi-Fi requirement).
- Telegram is already installed on most phones.
- Can forward photos easily if taken by someone else.
- Native multi-photo send in Telegram.

### Cons
- Requires internet connection on both phone and PC.
- Requires a Telegram account and a bot API token to be configured.
- Slight complexity in polling logic and handling different message types.
- Photos are stored on Telegram's servers temporarily (minor privacy consideration for invoice data).

### Dependencies
```
python-telegram-bot>=20.0
# or just: requests (already installed) — the bot API is a simple REST API
```

---

## Option D — Windows Phone Link (Microsoft)

### How it works
Microsoft's built-in "Phone Link" app (previously "Your Phone") pairs an Android phone to Windows and can share photos directly to the Windows clipboard or "recent photos" panel.

### Why it's not ideal for this use case
- iOS is **not supported** — Android only.
- Integration with a custom Python app is not straightforward (there's no API to hook into).
- Staff would still need to manually transfer photos from Phone Link into the app.
- It's a system-level app, not something the Python app can trigger or control.

---

## Option E — Email to Watched Inbox

### How it works
1. Create a dedicated email address (e.g. `invoices@yourshop.com` or a Gmail alias).
2. Staff email photos to that address from their phone.
3. App polls the inbox via IMAP and downloads attachments.

### Pros
- Works from anywhere, any phone, any email client.
- No special setup on the phone.

### Cons
- Email delivery delay (usually a few seconds, sometimes minutes).
- Requires IMAP credentials configured in the app.
- More complex to implement than the HTTP server option.
- Clutters an inbox with invoice photos.

---

## Option F — FTP/SFTP Server on PC

Similar in concept to Option A but using an FTP server app on the PC and an FTP client app on the phone. More setup overhead, no QR code convenience, and requires an FTP client app on the phone. Not recommended.

---

## Comparison Summary

| Option | Setup | Works off-LAN | Multi-photo | Phone app needed | Internet needed | Complexity |
|--------|-------|--------------|-------------|-----------------|----------------|-----------|
| A: HTTP+QR | Minimal | No | Yes | No | No | Low |
| B: Cloud sync | Cloud account | Yes | Yes | No (built-in sync) | Yes | Low |
| C: Telegram | Bot token | Yes | Yes | Yes (Telegram) | Yes | Medium |
| D: Phone Link | Windows pairing | No | No | No (Android only) | No | Not viable |
| E: Email/IMAP | Email account | Yes | Yes | No | Yes | Medium |

---

## Recommendation

**Start with Option A** (HTTP server + QR code). It requires the least external setup, works entirely on-premises, uses only one small new library (`qrcode[pil]`), and integrates directly with the existing image parse pipeline. The implementation is self-contained and can be added as a single new "Scan with Phone" button without touching any other code paths.

**Consider adding Option B as a companion** for situations where photos are taken outside the shop (e.g. at a trade show), letting them sync overnight and be processed the next morning.

---

## Notes for Implementation (Option A)

- The HTTP server should run in a daemon thread and be shut down when the dialog is closed (or after a timeout, e.g. 5 minutes).
- The QR code should display the PC's **local LAN IP** (e.g. `192.168.x.x`), not `localhost` or `127.0.0.1` (those won't work from the phone). Use `socket.gethostbyname(socket.gethostname())` or iterate `socket.getaddrinfo` to find the LAN IP.
- The upload endpoint should accept a `multipart/form-data` POST and return a simple "Upload received — you can close this page" response, so the user knows it worked.
- The temp images should be placed in the same temp directory pipeline already used by `_image_to_temp_pdf`.
- A "waiting for photos…" spinner in the dialog (using `CTkProgressBar` in indeterminate mode) gives visual feedback while the server waits.
- Consider a "Done" button on the phone page that sends a completion signal so the app knows all photos for a session have been uploaded (vs. a fixed timeout).
