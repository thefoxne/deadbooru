

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


