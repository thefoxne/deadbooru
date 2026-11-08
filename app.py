import os
import sys

# Ensure UTF-8 output on all Windows consoles (prevent cp932/cp1252 encode errors)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import csv
import glob
import json
import io
import re
import time
import base64
import argparse
import webbrowser
import threading
from typing import List, Dict, Tuple, Optional, Any
from collections import defaultdict

import requests
from flask import Flask, request, jsonify, render_template, Response

# ---------------------------------------------------------------------------
# Paths & Settings Manager (settings.json decoupled)
# ---------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

os.makedirs(DATA_DIR, exist_ok=True)

DEFAULT_SETTINGS = {
    "tags_file_path": "./data/tags.csv",
    "webhooks": [
        {
            "name": "Default Discord Channel",
            "url": ""
        }
    ],
    "active_webhook_name": "Default Discord Channel",
    "upload_mode": "discord_attachment",
    "auto_rating_tag": True,
    "default_rating": "general",
    "port": 5000
}

class SettingsManager:
    @staticmethod
    def load() -> Dict:
        if not os.path.exists(SETTINGS_FILE):
            SettingsManager.save(DEFAULT_SETTINGS)
            return DEFAULT_SETTINGS.copy()
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in DEFAULT_SETTINGS.items():
                    if k not in data:
                        data[k] = v
                return data
        except Exception as e:
            print(f"[!] Error loading {SETTINGS_FILE}: {e}")
            return DEFAULT_SETTINGS.copy()

    @staticmethod
    def save(data: Dict):
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"[+] Settings saved to {SETTINGS_FILE}")
            return True
        except Exception as e:
            print(f"[!] Error saving settings: {e}")
            return False

    @staticmethod
    def get_webhook_url(settings: Dict, webhook_name_or_url: Optional[str] = None) -> Optional[str]:
        webhooks = settings.get("webhooks", [])
        if not webhooks:
            return None

        if webhook_name_or_url and webhook_name_or_url.startswith("http"):
            return webhook_name_or_url

        target_name = webhook_name_or_url or settings.get("active_webhook_name")
        for wh in webhooks:
            if wh.get("name") == target_name and wh.get("url"):
                return wh.get("url").strip()

        for wh in webhooks:
            if wh.get("url"):
                return wh.get("url").strip()

        return None


# ---------------------------------------------------------------------------
# Decoupled Tag Database (Supports CSV and Excel .xlsx/.xls, In-Memory Only)
# ---------------------------------------------------------------------------

CATEGORY_MAP = {
    "0": "general",
    "1": "artist",
    "3": "copyright",
    "4": "character",
    "5": "meta",
    "art": "artist",
    "artist": "artist",
    "copy": "copyright",
    "copyright": "copyright",
    "series": "copyright",
    "char": "character",
    "character": "character",
    "general": "general",
    "meta": "meta"
}

class TagDatabase:
    def __init__(self, initial_path: Optional[str] = None):
        self.tags: List[Tuple[str, str, int]] = []
        self.tag_dict: Dict[str, Tuple[str, int]] = {}
        self.prefix_index: Dict[str, List[Tuple[str, str, int]]] = defaultdict(list)
        self.current_source: str = "None"
        self.current_filename: str = ""

        resolved_path = self.resolve_file(initial_path)
        if resolved_path:
            self.load_from_path(resolved_path)

    def resolve_file(self, requested_path: Optional[str] = None) -> Optional[str]:
        """Dynamically finds the tag data file instead of hardcoding."""
        candidates = []

        if requested_path:
            # Handle relative path from BASE_DIR
            p = requested_path if os.path.isabs(requested_path) else os.path.join(BASE_DIR, requested_path)
            if os.path.exists(p):
                return p
            candidates.append(p)

        # Check inside data/ directory for any .csv, .xlsx, .tsv (newest first)
        data_files = glob.glob(os.path.join(DATA_DIR, "*.*"))
        valid_files = [
            f for f in data_files 
            if f.lower().endswith((".csv", ".xlsx", ".xls", ".tsv"))
        ]
        if valid_files:
            valid_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return valid_files[0]

        # Fallback to Downloads directory
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        dl_tags = glob.glob(os.path.join(downloads_dir, "tags*.csv"))
        if dl_tags:
            dl_tags.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return dl_tags[0]

        return None

    def load_from_path(self, file_path: str) -> bool:
        if not os.path.exists(file_path):
            print(f"[!] File not found: {file_path}")
            return False

        ext = os.path.splitext(file_path)[1].lower()
        print(f"[*] Loading tag data from: {file_path} (Format: {ext})...")

        raw_items = []
        try:
            if ext in (".xlsx", ".xls"):
                raw_items = self._load_excel(file_path)
            else:
                raw_items = self._load_csv(file_path)

            self._build_index(raw_items, file_path)
            print(f"[+] Loaded {len(self.tags):,} tags into memory from {os.path.basename(file_path)}.")
            return True
        except Exception as e:
            print(f"[!] Error loading {file_path}: {e}")
            return False

    def load_from_stream(self, stream_bytes: bytes, filename: str) -> bool:
        """Loads tag data directly from memory stream (uploaded via UI)."""
        ext = os.path.splitext(filename)[1].lower()
        try:
            if ext in (".xlsx", ".xls"):
                raw_items = self._load_excel_stream(stream_bytes)
            else:
                raw_items = self._load_csv_stream(stream_bytes)

            self._build_index(raw_items, filename)
            print(f"[+] Uploaded and loaded {len(self.tags):,} tags into memory from {filename}.")
            return True
        except Exception as e:
            print(f"[!] Error loading stream {filename}: {e}")
            return False

    def _load_csv(self, file_path: str) -> List[Tuple[str, str, int]]:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            return self._parse_csv_lines(f)

    def _load_csv_stream(self, stream_bytes: bytes) -> List[Tuple[str, str, int]]:
        text = stream_bytes.decode("utf-8", errors="replace")
        return self._parse_csv_lines(io.StringIO(text))

    def _parse_csv_lines(self, f_obj) -> List[Tuple[str, str, int]]:
        reader = csv.reader(f_obj)
        header = next(reader, None)
        if not header:
            return []

        col_map = {str(col).strip().lower(): idx for idx, col in enumerate(header) if col}
        name_idx = col_map.get("name", col_map.get("tag", col_map.get("tag_name", 0)))
        cat_idx = col_map.get("category", col_map.get("cat", col_map.get("type", 1 if len(header) > 1 else -1)))
        cnt_idx = col_map.get("count", col_map.get("posts", col_map.get("post_count", 2 if len(header) > 2 else -1)))

        items = []
        for row in reader:
            if not row or len(row) <= name_idx:
                continue
            name = row[name_idx].strip()
            if not name:
                continue

            raw_cat = row[cat_idx].strip().lower() if cat_idx >= 0 and len(row) > cat_idx else "general"
            cat = CATEGORY_MAP.get(raw_cat, "general")

            raw_cnt = row[cnt_idx].strip() if cnt_idx >= 0 and len(row) > cnt_idx else "0"
            try:
                cnt = int(raw_cnt) if raw_cnt else 0
            except (ValueError, TypeError):
                cnt = 0

            items.append((name, cat, cnt))
        return items

    def _load_excel(self, file_path: str) -> List[Tuple[str, str, int]]:
        try:
            import openpyxl
        except ImportError:
            print("[!] openpyxl is not installed. Please run `pip install openpyxl` to support Excel files.")
            return []

        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        return self._parse_excel_workbook(wb)

    def _load_excel_stream(self, stream_bytes: bytes) -> List[Tuple[str, str, int]]:
        try:
            import openpyxl
        except ImportError:
            return []

        wb = openpyxl.load_workbook(io.BytesIO(stream_bytes), read_only=True, data_only=True)
        return self._parse_excel_workbook(wb)

    def _parse_excel_workbook(self, wb) -> List[Tuple[str, str, int]]:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            wb.close()
            return []

        col_map = {str(col).strip().lower(): idx for idx, col in enumerate(header) if col is not None}
        name_idx = col_map.get("name", col_map.get("tag", col_map.get("tag_name", 0)))
        cat_idx = col_map.get("category", col_map.get("cat", col_map.get("type", 1 if len(header) > 1 else -1)))
        cnt_idx = col_map.get("count", col_map.get("posts", col_map.get("post_count", 2 if len(header) > 2 else -1)))

        items = []
        for r in rows:
            if not r or len(r) <= name_idx or not r[name_idx]:
                continue
            name = str(r[name_idx]).strip()
            if not name:
                continue

            raw_cat = str(r[cat_idx]).strip().lower() if cat_idx >= 0 and len(r) > cat_idx and r[cat_idx] else "general"
            cat = CATEGORY_MAP.get(raw_cat, "general")

            raw_cnt = r[cnt_idx] if cnt_idx >= 0 and len(r) > cnt_idx else 0
            try:
                cnt = int(raw_cnt) if raw_cnt is not None else 0
            except (ValueError, TypeError):
                cnt = 0

            items.append((name, cat, cnt))
        wb.close()
        return items

    def _build_index(self, items: List[Tuple[str, str, int]], source_name: str):
        self.tags = items
        self.tags.sort(key=lambda x: x[2], reverse=True)
        self.tag_dict = {item[0]: (item[1], item[2]) for item in self.tags}

        self.prefix_index = defaultdict(list)
        for item in self.tags:
            name = item[0]
            if len(name) >= 1:
                self.prefix_index[name[:1].lower()].append(item)
            if len(name) >= 2:
                self.prefix_index[name[:2].lower()].append(item)

        self.current_source = source_name
        self.current_filename = os.path.basename(source_name)

    def search(self, query: str, limit: int = 15, category: Optional[str] = None) -> List[Dict]:
        query = query.lower().strip().replace(" ", "_")
        if not query:
            return []

        is_wildcard = query.startswith("*")
        clean_q = query.lstrip("*") if is_wildcard else query
        if not clean_q:
            return []

        prefix_matches: List[Tuple[str, str, int]] = []
        sub_matches: List[Tuple[str, str, int]] = []

        search_candidates = self.tags
        if not is_wildcard and len(clean_q) in (1, 2) and clean_q in self.prefix_index:
            search_candidates = self.prefix_index[clean_q]

        for item in search_candidates:
            name, cat, cnt = item
            if category and category != "all" and cat != category:
                continue

            name_lower = name.lower()

            if not is_wildcard and name_lower.startswith(clean_q):
                prefix_matches.append(item)
                if len(prefix_matches) >= limit:
                    break
            elif clean_q in name_lower:
                if len(sub_matches) < limit:
                    sub_matches.append(item)

        results = []
        seen = set()

        for item in prefix_matches + sub_matches:
            name, cat, cnt = item
            if name not in seen:
                seen.add(name)
                results.append({
                    "name": name,
                    "category": cat,
                    "count": cnt,
                    "count_formatted": self.format_count(cnt)
                })
                if len(results) >= limit:
                    break

        return results

    def get_info(self, tag_names: List[str]) -> Dict[str, Dict]:
        res = {}
        for tag in tag_names:
            clean_tag = tag.lower().strip()
            if clean_tag in self.tag_dict:
                cat, cnt = self.tag_dict[clean_tag]
                res[clean_tag] = {
                    "category": cat,
                    "count": cnt,
                    "count_formatted": self.format_count(cnt),
                    "known": True
                }
            else:
                inferred_cat = "general"
                if ":" in clean_tag:
                    prefix = clean_tag.split(":", 1)[0]
                    if prefix in ("art", "artist"):
                        inferred_cat = "artist"
                    elif prefix in ("copy", "copyright"):
                        inferred_cat = "copyright"
                    elif prefix in ("char", "character"):
                        inferred_cat = "character"
                    elif prefix in ("rating", "meta", "status", "source", "pool"):
                        inferred_cat = "meta"

                res[clean_tag] = {
                    "category": inferred_cat,
                    "count": 0,
                    "count_formatted": "0",
                    "known": False
                }
        return res

    @staticmethod
    def format_count(count: int) -> str:
        if count >= 1_000_000:
            return f"{count / 1_000_000:.1f}M"
        if count >= 1_000:
            return f"{count / 1_000:.1f}k"
        return str(count)


# ---------------------------------------------------------------------------
# Flask Application Setup
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

tag_db: TagDatabase = TagDatabase()
session_posts: List[Dict] = []
post_counter: int = 1000


@app.route("/")
def index():
    settings = SettingsManager.load()
    tag_count = f"{len(tag_db.tags):,}" if tag_db else "0"
    loaded_file = tag_db.current_filename if tag_db else "None"
    return render_template(
        "index.html",
        tag_count=tag_count,
        loaded_file=loaded_file,
        settings=settings
    )


# ---------------------------------------------------------------------------
# Settings & Tag File Management API Endpoints
# ---------------------------------------------------------------------------

@app.route("/api/settings", methods=["GET"])
def api_get_settings():
    return jsonify(SettingsManager.load())


@app.route("/api/settings", methods=["POST"])
def api_save_settings():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return jsonify({"error": "Invalid data format"}), 400

    success = SettingsManager.save(data)
    if success:
        return jsonify({"success": True, "settings": data})
    else:
        return jsonify({"error": "Failed to save settings.json"}), 500


@app.route("/api/tags/status", methods=["GET"])
def api_tags_status():
    """Returns current loaded tag file status and list of files in ./data/ folder."""
    data_files = []
    if os.path.exists(DATA_DIR):
        for f in os.listdir(DATA_DIR):
            if f.lower().endswith((".csv", ".xlsx", ".xls", ".tsv")):
                full_p = os.path.join(DATA_DIR, f)
                data_files.append({
                    "name": f,
                    "path": f"./data/{f}",
                    "size": f"{os.path.getsize(full_p) / (1024*1024):.2f} MB",
                    "modified": time.strftime("%H:%M %d/%m/%Y", time.localtime(os.path.getmtime(full_p)))
                })

    return jsonify({
        "current_source": tag_db.current_source if tag_db else "",
        "current_filename": tag_db.current_filename if tag_db else "",
        "tag_count": len(tag_db.tags) if tag_db else 0,
        "tag_count_formatted": f"{len(tag_db.tags):,}" if tag_db else "0",
        "available_files": data_files
    })


@app.route("/api/tags/reload", methods=["POST"])
def api_tags_reload():
    """
    Reloads tag file either:
    1. From a specified path (JSON { "path": "..." })
    2. Or from an uploaded file (multipart/form-data 'file')
    """
    global tag_db

    # Check if a file was uploaded directly through UI
    if "file" in request.files:
        uploaded_file = request.files["file"]
        if uploaded_file.filename:
            filename = uploaded_file.filename
            file_bytes = uploaded_file.read()

            # Save a copy to ./data/ folder so it persists for future runs
            save_path = os.path.join(DATA_DIR, filename)
            try:
                with open(save_path, "wb") as f:
                    f.write(file_bytes)
            except Exception as e:
                print(f"[!] Warning: could not write to {save_path}: {e}")

            # Load into RAM
            ok = tag_db.load_from_stream(file_bytes, filename)
            if ok:
                # Update settings.json to remember this file
                settings = SettingsManager.load()
                settings["tags_file_path"] = f"./data/{filename}"
                SettingsManager.save(settings)

                return jsonify({
                    "success": True,
                    "filename": filename,
                    "tag_count": len(tag_db.tags),
                    "tag_count_formatted": f"{len(tag_db.tags):,}"
                })
            else:
                return jsonify({"error": f"Unable to read file {filename}"}), 400

    # Otherwise reload from path
    data = request.get_json(silent=True) or {}
    req_path = data.get("path")
    if not req_path:
        settings = SettingsManager.load()
        req_path = settings.get("tags_file_path")

    resolved = tag_db.resolve_file(req_path)
    if not resolved:
        return jsonify({"error": f"Tag file not found at: {req_path}"}), 400

    ok = tag_db.load_from_path(resolved)
    if ok:
        settings = SettingsManager.load()
        settings["tags_file_path"] = req_path
        SettingsManager.save(settings)

        return jsonify({
            "success": True,
            "filename": os.path.basename(resolved),
            "tag_count": len(tag_db.tags),
            "tag_count_formatted": f"{len(tag_db.tags):,}"
        })
    else:
        return jsonify({"error": f"Error reading file {resolved}"}), 500


# ---------------------------------------------------------------------------
# Posts Hub API Endpoints (In-Memory Session History)
# ---------------------------------------------------------------------------

@app.route("/api/posts", methods=["GET"])
def api_get_posts():
    return jsonify(session_posts)


# ---------------------------------------------------------------------------
# Tags & Search API Endpoints
# ---------------------------------------------------------------------------

@app.route("/api/tags/autocomplete", methods=["GET"])
def api_autocomplete():
    q = request.args.get("q", "").strip()
    try:
        limit = int(request.args.get("limit", 15))
    except (ValueError, TypeError):
        limit = 15

    if not tag_db:
        return jsonify([])

    results = tag_db.search(q, limit=limit)
    return jsonify(results)


@app.route("/api/tags/search", methods=["GET"])
def api_tags_search():
    """Full search endpoint for Danbooru Tags Hub."""
    q = request.args.get("q", "").strip()
    cat = request.args.get("category", "all").strip().lower()
    try:
        limit = int(request.args.get("limit", 40))
    except (ValueError, TypeError):
        limit = 40

    if not tag_db:
        return jsonify([])

    if not q:
        res = []
        for item in tag_db.tags:
            if cat == "all" or item[1] == cat:
                res.append({
                    "name": item[0],
                    "category": item[1],
                    "count": item[2],
                    "count_formatted": tag_db.format_count(item[2])
                })
                if len(res) >= limit:
                    break
        return jsonify(res)

    results = tag_db.search(q, limit=limit, category=cat)
    return jsonify(results)


@app.route("/api/tags/batch", methods=["POST"])
def api_tags_batch():
    data = request.get_json(silent=True) or {}
    tags = data.get("tags", [])
    if not isinstance(tags, list):
        return jsonify({})

    if not tag_db:
        return jsonify({})

    info = tag_db.get_info(tags)
    return jsonify(info)


@app.route("/api/proxy_image")
def api_proxy_image():
    """Stream image from Pixiv or Twitter with correct Referer/User-Agent directly into RAM."""
    raw_url = request.args.get("url", "").strip()
    if not raw_url:
        return "Missing URL", 400

    referer = "https://www.pixiv.net/" if "pximg.net" in raw_url else raw_url
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": referer
    }
    try:
        r = requests.get(raw_url, headers=headers, timeout=15)
        if r.status_code == 200:
            content_type = r.headers.get("Content-Type", "image/jpeg").split(";")[0].strip()
            return Response(r.content, mimetype=content_type)
        return f"Proxy failed: HTTP {r.status_code}", 502
    except Exception as e:
        return str(e), 500


def resolve_pixiv(url: str) -> Optional[Dict[str, Any]]:
    """Resolves Pixiv artwork including multi-page posts, metadata, and tags."""
    m = re.search(r'pixiv\.net/(?:en/)?artworks/(\d+)', url)
    if not m:
        return None
    illust_id = m.group(1)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.pixiv.net/"
    }
    try:
        r = requests.get(f"https://www.pixiv.net/ajax/illust/{illust_id}", headers=headers, timeout=8)
        if r.status_code != 200:
            return None
        body = r.json().get("body", {})
        artist = body.get("userName", "")
        title = body.get("title", "")
        tags = [t["tag"] for t in body.get("tags", {}).get("tags", [])]
        published = body.get("uploadDate", "")

        # Check pages endpoint
        r_pages = requests.get(f"https://www.pixiv.net/ajax/illust/{illust_id}/pages", headers=headers, timeout=8)
        pages = []
        if r_pages.status_code == 200:
            for i, p in enumerate(r_pages.json().get("body", [])):
                urls = p.get("urls", {})
                reg = urls.get("regular") or urls.get("original")
                thumb = urls.get("thumb_mini") or urls.get("small") or reg
                pages.append({
                    "index": i,
                    "page_num": i + 1,
                    "width": p.get("width", 0),
                    "height": p.get("height", 0),
                    "img_url": reg,
                    "thumb_url": thumb,
                    "source_url": f"https://www.pixiv.net/artworks/{illust_id}#{i+1}"
                })

        if not pages:
            reg = body.get("urls", {}).get("regular") or body.get("urls", {}).get("original")
            if reg:
                pages.append({
                    "index": 0,
                    "page_num": 1,
                    "width": body.get("width", 0),
                    "height": body.get("height", 0),
                    "img_url": reg,
                    "thumb_url": body.get("urls", {}).get("thumb") or reg,
                    "source_url": f"https://www.pixiv.net/artworks/{illust_id}"
                })

        if not pages:
            return None

        first_img_url = pages[0]["img_url"]
        img_r = requests.get(first_img_url, headers=headers, timeout=12)
        if img_r.status_code != 200:
            return None
        ct = img_r.headers.get("Content-Type", "image/jpeg").split(";")[0].strip().lower()
        b64 = base64.b64encode(img_r.content).decode("utf-8")
        data_uri = f"data:{ct};base64,{b64}"

        return {
            "success": True,
            "type": "pixiv",
            "is_multi": len(pages) > 1,
            "count": len(pages),
            "current_index": 0,
            "artist": artist,
            "title": title,
            "tags": tags,
            "published": published,
            "pages": pages,
            "data_uri": data_uri,
            "size": len(img_r.content),
            "content_type": ct,
            "direct_url": pages[0]["img_url"],
            "source_url": pages[0]["source_url"]
        }
    except Exception as e:
        print(f"[!] resolve_pixiv error: {e}")
        return None


def resolve_twitter(url: str) -> Optional[Dict[str, Any]]:
    """Resolves Twitter artwork including multi-photo tweets, author, and media."""
    m = re.search(r'(?:twitter\.com|x\.com)/([^/]+)/status/(\d+)', url)
    if not m:
        return None
    user, tweet_id = m.group(1), m.group(2)
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        r = requests.get(f"https://api.fxtwitter.com/{user}/status/{tweet_id}", headers=headers, timeout=8)
        if r.status_code != 200:
            return None
        tweet = r.json().get("tweet", {})
        artist = tweet.get("author", {}).get("name", "")
        screen_name = tweet.get("author", {}).get("screen_name", user)
        published = tweet.get("created_at", "")
        photos = tweet.get("media", {}).get("photos", [])
        if not photos:
            return None

        pages = []
        for i, p in enumerate(photos):
            img_url = p.get("url")
            pages.append({
                "index": i,
                "page_num": i + 1,
                "width": p.get("width", 0),
                "height": p.get("height", 0),
                "img_url": img_url,
                "thumb_url": img_url,
                "source_url": f"https://x.com/{screen_name}/status/{tweet_id}/photo/{i+1}"
            })

        first_img_url = pages[0]["img_url"]
        img_r = requests.get(first_img_url, headers=headers, timeout=12)
        if img_r.status_code != 200:
            return None
        ct = img_r.headers.get("Content-Type", "image/jpeg").split(";")[0].strip().lower()
        b64 = base64.b64encode(img_r.content).decode("utf-8")
        data_uri = f"data:{ct};base64,{b64}"

        return {
            "success": True,
            "type": "twitter",
            "is_multi": len(pages) > 1,
            "count": len(pages),
            "current_index": 0,
            "artist": artist,
            "screen_name": screen_name,
            "title": tweet.get("text", "")[:60],
            "tags": [],
            "published": published,
            "pages": pages,
            "data_uri": data_uri,
            "size": len(img_r.content),
            "content_type": ct,
            "direct_url": pages[0]["img_url"],
            "source_url": pages[0]["source_url"]
        }
    except Exception as e:
        print(f"[!] resolve_twitter error: {e}")
        return None


@app.route("/api/fetch_image", methods=["POST"])
def api_fetch_image():
    """Fetch an image or multi-image post from URL directly into RAM without saving to disk."""
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "No URL provided"}), 400

    # 1. Specialized Twitter/X artwork resolver (supports multi-photo)
    if "twitter.com" in url or "x.com" in url:
        tw_res = resolve_twitter(url)
        if tw_res:
            return jsonify(tw_res)

    # 2. Specialized Pixiv artwork resolver (supports multi-page)
    if "pixiv.net" in url:
        px_res = resolve_pixiv(url)
        if px_res:
            return jsonify(px_res)

    # 3. Direct image or generic OpenGraph check
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": url
    }

    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code != 200:
            return jsonify({"error": f"Failed to fetch image: HTTP {resp.status_code}"}), 400

        content_type = resp.headers.get("Content-Type", "image/png").split(";")[0].strip().lower()
        if content_type.startswith("image/"):
            b64_data = base64.b64encode(resp.content).decode("utf-8")
            data_uri = f"data:{content_type};base64,{b64_data}"
            return jsonify({
                "success": True,
                "is_multi": False,
                "count": 1,
                "data_uri": data_uri,
                "size": len(resp.content),
                "content_type": content_type,
                "direct_url": url,
                "source_url": url
            })

        # Generic OpenGraph check (STRICTLY ignoring profile avatars)
        html_text = resp.text
        og_match = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']', html_text, re.IGNORECASE)
        if not og_match:
            og_match = re.search(r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image["\']', html_text, re.IGNORECASE)
        if not og_match:
            og_match = re.search(r'<meta[^>]+name=["\']twitter:image["\'][^>]+content=["\']([^"\']+)["\']', html_text, re.IGNORECASE)

        if og_match:
            og_img_url = og_match.group(1).replace("&amp;", "&")
            og_lower = og_img_url.lower()
            # NEVER accept user avatar or profile images as artwork!
            if not any(bad in og_lower for bad in ("profile_images", "avatar", "default_profile", "user_avatar")):
                try:
                    img_resp = requests.get(og_img_url, headers=headers, timeout=10)
                    if img_resp.status_code == 200 and img_resp.headers.get("Content-Type", "").startswith("image/"):
                        img_ct = img_resp.headers.get("Content-Type", "image/png").split(";")[0].strip().lower()
                        b64_data = base64.b64encode(img_resp.content).decode("utf-8")
                        return jsonify({
                            "success": True,
                            "is_multi": False,
                            "count": 1,
                            "data_uri": f"data:{img_ct};base64,{b64_data}",
                            "size": len(img_resp.content),
                            "content_type": img_ct,
                            "direct_url": og_img_url,
                            "source_url": url
                        })
                except Exception:
                    pass

        return jsonify({"error": "URL does not return a direct image", "is_page": True}), 400

    except Exception as e:
        return jsonify({"error": f"Exception fetching image: {str(e)}"}), 500


@app.route("/api/test_webhook", methods=["POST"])
def api_test_webhook():
    """Test if Discord webhook URL is reachable."""
    data = request.get_json(silent=True) or {}
    webhook_url = data.get("webhook_url", "").strip()

    if not webhook_url:
        settings = SettingsManager.load()
        webhook_name = data.get("webhook_name")
        webhook_url = SettingsManager.get_webhook_url(settings, webhook_name)

    if not webhook_url:
        return jsonify({"error": "Discord Webhook URL not found!"}), 400

    try:
        resp = requests.get(webhook_url, timeout=10)
        if resp.status_code == 200:
            info = resp.json()
            return jsonify({
                "success": True,
                "name": info.get("name", "Webhook"),
                "channel_id": info.get("channel_id")
            })
        else:
            return jsonify({
                "success": False,
                "error": f"Discord responded with error code {resp.status_code}: {resp.text}"
            }), 400
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


# -----------------------------------------------------------
# Upload / Post Action
# -----------------------------------------------------------

@app.route("/api/upload", methods=["POST"])
def api_upload():
    global post_counter, session_posts

    settings = SettingsManager.load()

    webhook_name = request.form.get("webhook_name", "").strip()
    webhook_url = request.form.get("webhook_url", "").strip()
    if not webhook_url:
        webhook_url = SettingsManager.get_webhook_url(settings, webhook_name)

    if not webhook_url:
        return jsonify({"error": "Please configure a Discord Webhook URL in Settings before posting!"}), 400

    source_link = request.form.get("source_link", "").strip()
    tags_text = request.form.get("tags", "").strip()

    # Normalize tags into clean lowercase single-space separated format (filtering out accidental URLs)
    tags_list = [
        t.strip().lower() 
        for t in re.split(r"\s+", tags_text) 
        if t.strip() and not t.strip().lower().startswith(("http://", "https://"))
    ]
    formatted_tags = " ".join(tags_list)

    file = request.files.get("file")
    effective_link = source_link

    # If no source_link provided but user uploaded a file, host to Catbox to get public URL
    if not effective_link and file and file.filename:
        try:
            filename = file.filename
            content_type = file.content_type or "image/png"
            file_bytes = file.read()
            catbox_resp = requests.post(
                "https://catbox.moe/user/api.php",
                data={"reqtype": "fileupload"},
                files={"fileToUpload": (filename, io.BytesIO(file_bytes), content_type)},
                timeout=30
            )
            if catbox_resp.status_code == 200 and catbox_resp.text.startswith("http"):
                effective_link = catbox_resp.text.strip()
        except Exception as e:
            print(f"[!] Catbox upload fallback error: {e}")

    if not effective_link:
        return jsonify({"error": "Please enter a source URL or post link (Twitter, Pixiv...)!"}), 400

    # Build exact requested Danbooru upload URL format:
    # https://danbooru.donmai.us/uploads/new?url=link.com
    # ` taga tagb tagc ... `
    danbooru_upload_url = f"https://danbooru.donmai.us/uploads/new?url={effective_link}"
    tags_line = f"` {formatted_tags} `"
    discord_content = f"{danbooru_upload_url}\n{tags_line}"

    # Pure text to Discord webhook (NO IMAGE ATTACHMENT)
    try:
        resp = requests.post(
            webhook_url,
            json={"content": discord_content},
            timeout=15
        )

        if resp.status_code not in (200, 204):
            return jsonify({"error": f"Discord Webhook error ({resp.status_code}): {resp.text}"}), 400

        resp_data = resp.json() if resp.text else {}
        message_id = resp_data.get("id")

        post_counter += 1
        post_record = {
            "id": post_counter,
            "link": danbooru_upload_url,
            "source_link": effective_link,
            "tags": tags_line,
            "tags_list": tags_list,
            "thumbnail": None,
            "message_id": message_id,
            "timestamp": time.strftime("%H:%M:%S - %d/%m/%Y")
        }
        session_posts.insert(0, post_record)
        if len(session_posts) > 50:
            session_posts.pop()

        return jsonify({
            "success": True,
            "message": "Posted to Discord successfully!",
            "content": discord_content,
            "danbooru_url": danbooru_upload_url,
            "source_link": effective_link,
            "tags": tags_line,
            "message_id": message_id,
            "post": post_record
        })

    except Exception as e:
        return jsonify({"error": f"Webhook delivery error: {str(e)}"}), 500


# ---------------------------------------------------------------------------
# CLI & Server Runner
# ---------------------------------------------------------------------------

def find_available_port(start_port: int) -> int:
    import socket
    port = start_port
    while port < start_port + 50:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            res = s.connect_ex(("127.0.0.1", port))
            if res != 0:
                return port
        port += 1
    return start_port

def open_browser(url: str):
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()

def main():
    global tag_db

    settings = SettingsManager.load()

    parser = argparse.ArgumentParser(description="Danbooru Discord Offline Hub & Uploader")
    parser.add_argument("--port", type=int, default=settings.get("port", 5000), help="Port to run localhost server on")
    parser.add_argument("--tags-file", type=str, default=settings.get("tags_file_path"), help="Path to tags CSV/Excel file")
    parser.add_argument("--no-browser", action="store_true", help="Do not automatically open browser on launch")

    args = parser.parse_args()

    # Load tag database without hardcoding
    tag_db = TagDatabase(args.tags_file)

    port = find_available_port(args.port)
    local_url = f"http://127.0.0.1:{port}"
    print("=" * 68)
    print(f"  DANBOORU DISCORD HUB & UPLOADER - LOCALHOST SERVER")
    print(f"  URL: {local_url}")
    print(f"  Settings: {SETTINGS_FILE}")
    print(f"  Tag Data: {tag_db.current_source} ({len(tag_db.tags):,} tags)")
    print(f"  Memory mode: Zero local disk storage for posts/images")
    print("=" * 68)

    if not args.no_browser:
        open_browser(local_url)

    app.run(host="127.0.0.1", port=port, debug=False)

if __name__ == "__main__":
    main()
