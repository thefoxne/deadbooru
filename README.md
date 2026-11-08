# deadbooru

A lightweight, standalone offline Danbooru tagging and upload hub with Discord webhook integration. Designed to keep Danbooru tagging workflows active with high efficiency.

---

## Features

### 1. Pure-Text Danbooru Payload for Discord
- Posts pure-text payloads directly to Discord channels via Webhooks:
  ```text
  https://danbooru.donmai.us/uploads/new?url=<source_url>
  ` taga tagb tagc ... `
  ```
- **Line 1:** Direct Danbooru upload initialization URL preloaded with the source image/artwork link.
- **Line 2:** Clean, lowercase tags wrapped inside a monospace code block (`` ` ... ` ``) to avoid Discord markdown formatting interference and allow instant one-click copying.

### 2. Multi-Image & Album Support
- Automatically detects and resolves multi-image posts (Pixiv albums, Twitter/X multi-photo tweets).
- Gallery view matching Danbooru's real thumbnail card layout:
  - Artist name with Danbooru quick links
  - Clickable tag chips to add tags directly
  - Filtering by **All**, **Posted**, and **Unposted**
  - Per-image tagging and posting status tracking

### 3. Smooth Zoom & Pan (Up to 1000%)
- Inspect fine details with smooth mouse wheel zooming from **50% up to 1000% (10x)** centered on the mouse cursor.
- Click and drag to pan around enlarged images with zero ghost drag interference.
- Quick toolbar controls: Zoom Out, Fit to frame, 1:1 original size, Zoom In, and 90° image rotation.

### 4. 72,000+ Danbooru Tags in RAM
- In-memory tag database loaded from decoupled `./data/tags.csv` (or custom `.xlsx` / `.csv` spreadsheets).
- Fast real-time autocomplete suggestions with Danbooru category colors (Artist red, Copyright purple, Character green, General blue, Meta orange).
- Tag browser tab for searching and exploring tags offline.

### 5. Zero Local Disk Storage
- Images and pasted clipboard contents are processed entirely in memory via byte streams (`io.BytesIO`) without creating temporary cache files or junk on disk.

### 6. Dynamic Webhook & Configuration Management
- Manage multiple Discord webhooks from an in-app overlay modal dialog (`settings.json`).
- Switch active webhooks or test connection status on the fly.

---

## Quick Start

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/thefoxne/deadbooru.git
   cd deadbooru
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Running

- **On Windows:** Double-click `run.bat` or run:
  ```bash
  python app.py
  ```
- **Custom Port or Tag File:**
  ```bash
  python app.py --port 5000 --tags-file ./data/tags.csv
  ```

Open your browser at `http://127.0.0.1:5000`.

---

## Configuration

Settings are saved to `settings.json`. An example configuration is provided in `settings.example.json`:

```json
{
  "active_webhook_name": "Default Discord Channel",
  "auto_rating_tag": true,
  "default_rating": "general",
  "port": 5000,
  "tags_file_path": "./data/tags.csv",
  "upload_mode": "text_only",
  "webhooks": [
    {
      "name": "Default Discord Channel",
      "url": "https://discord.com/api/webhooks/YOUR_WEBHOOK_ID/YOUR_WEBHOOK_TOKEN"
    }
  ]
}
```

---

## License

MIT License.
