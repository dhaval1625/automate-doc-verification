import os
import re
import sys
from typing import List, Optional, Tuple
from urllib.parse import urlparse
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Ensure UTF-8 output encoding across Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# --- SHARED CONFIGURATION ---
SPREADSHEET_ID = "10drtwlduGC1-5V1yNYFNtvBTfHM4Klre25Kzfbmirxw"
DEFAULT_SHEET_NAME = "Current Site Pages & New Site Pages"
DEFAULT_START_ROW = 313
LAST_ROW_TO_CHECK = 457
DEFAULT_BATCH_SIZE = 10

JEKYLL_PROJECT_DIR = r"D:\projects\document-verify\manektech-2026-jekyll"

CONTENT_FOLDERS = [
    "_services",
    "_technologies",
    "_solutions",
    "_industries",
    "_pages",
    "_work",
    "_blogposts",
    "_ebooks",
    "_events",
    "_podcast",
    "_tutorials",
    "_whitepapers",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/documents.readonly"
]

ROWS_TO_SKIP = [137, 138, 139, 140, 141]

# --- AGENT SESSION PERSISTENCE ---

def get_latest_agy_conversation_id() -> Optional[str]:
    """Finds the most recent conversation ID created by agy CLI."""
    conv_dir = os.path.expanduser(r"~/.gemini/antigravity-cli/conversations")
    if os.path.isdir(conv_dir):
        try:
            files = [
                os.path.join(conv_dir, f)
                for f in os.listdir(conv_dir)
                if f.endswith(".db") or f.endswith(".pb")
            ]
            if files:
                files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                base = os.path.basename(files[0])
                m = re.match(r"^([a-f0-9-]{36})", base)
                if m:
                    return m.group(1)
        except Exception:
            pass
    return None

# --- GOOGLE API SERVICES ---

def get_services():
    """Authenticates and initializes Google Sheets, Drive, and Docs client services."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    docs_service = build("docs", "v1", credentials=creds)
    return sheets_service, drive_service, docs_service

def get_resolved_sheet_name(sheets_service, spreadsheet_id: str, preferred: str = DEFAULT_SHEET_NAME) -> str:
    """Finds the actual tab title in the spreadsheet that matches the preferred name."""
    try:
        meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = meta.get("sheets", [])
        titles = [s.get("properties", {}).get("title", "") for s in sheets]

        # 1. Exact match
        for t in titles:
            if t == preferred:
                return t

        # 2. Case-insensitive match
        for t in titles:
            if t.lower() == preferred.lower():
                return t

        # 3. Substring match
        clean_preferred = preferred.lstrip("1234567890. ").strip().lower()
        for t in titles:
            if clean_preferred in t.lower():
                return t

        # 4. Fallback to first tab if none matched
        if titles:
            return titles[0]
    except Exception as e:
        print(f"⚠️ Could not inspect spreadsheet metadata: {e}")

    return preferred

def extract_doc_id(url: str) -> Optional[str]:
    """Extracts the Google Doc ID from a URL."""
    if not url:
        return None
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

# --- URL & JEKYLL FILE MATCHING ---

def normalize_url_path(url: str) -> str:
    """Extracts lowercase stripped path from a URL (e.g., 'https://domain.com/path/' -> '/path')."""
    if not url:
        return ""
    if url.startswith("http://") or url.startswith("https://"):
        parsed = urlparse(url)
        path = parsed.path.rstrip("/")
        return path.lower()
    path = url.split("?")[0].split("#")[0].rstrip("/")
    if not path.startswith("/") and not path.startswith("http"):
        path = "/" + path
    return path.lower()

def get_target_directories(jekyll_dir: str = JEKYLL_PROJECT_DIR, folders: Optional[List[str]] = None) -> List[str]:
    """Returns a list of existing directories to search for Jekyll content files."""
    active_folders = folders if folders is not None else CONTENT_FOLDERS
    target_dirs = []

    if active_folders:
        for folder in active_folders:
            folder_path = folder if os.path.isabs(folder) else os.path.join(jekyll_dir, folder)
            if os.path.isdir(folder_path):
                target_dirs.append(folder_path)

    # Fallback to base jekyll directory if none of the specific folders exist
    if not target_dirs and os.path.isdir(jekyll_dir):
        target_dirs.append(jekyll_dir)

    return target_dirs

def find_matching_jekyll_file(
    page_name: str,
    live_url: str,
    jekyll_dir: str = JEKYLL_PROJECT_DIR,
    folders: Optional[List[str]] = None
) -> Optional[str]:
    """
    Finds the local .md file across target Jekyll content folders based on:
    1. Direct file name match ({slug}.md)
    2. File name substring
    3. Front matter permalink or title
    """
    target_dirs = get_target_directories(jekyll_dir, folders)

    slug = ""
    if live_url:
        path = urlparse(live_url).path.rstrip("/")
        if path:
            slug = path.split("/")[-1].lower()

    page_name_slug = re.sub(r"[^a-zA-Z0-9]+", "-", page_name.lower()).strip("-") if page_name else ""
    search_slugs = [s for s in [slug, page_name_slug] if s]

    if not search_slugs:
        return None

    # 1. Direct filename match in any target directory
    for d in target_dirs:
        for s in search_slugs:
            candidate = os.path.join(d, f"{s}.md")
            if os.path.isfile(candidate):
                return candidate

    # 2. Walk target directories: check filename substring
    all_md_files = []
    ignored_subdirs = {"_site", ".jekyll-cache", ".git", "node_modules", "vendor", ".antigravity", "assets"}
    for d in target_dirs:
        for root, _, files in os.walk(d):
            if any(ignored in root for ignored in ignored_subdirs):
                continue
            for f in files:
                if f.endswith(".md"):
                    full_path = os.path.join(root, f)
                    all_md_files.append(full_path)
                    f_lower = f.lower()
                    for s in search_slugs:
                        if s in f_lower:
                            return full_path

    # 3. Check front matter permalink / title
    for fpath in all_md_files:
        try:
            with open(fpath, "r", encoding="utf-8", errors="ignore") as fp:
                head = fp.read(2048)

            # Check permalink in frontmatter
            m = re.search(r"permalink:\s*([^\s\n\r]+)", head, re.IGNORECASE)
            if m:
                permalink = m.group(1).strip().strip("\"'").rstrip("/").lower()
                for s in search_slugs:
                    if permalink.endswith(f"/{s}") or permalink == f"/{s}" or permalink == s:
                        return fpath

            # Check title in frontmatter
            m_title = re.search(r"^title:\s*[\"']?(.*?)[\"']?$", head, re.MULTILINE | re.IGNORECASE)
            if m_title and page_name:
                if m_title.group(1).strip().lower() == page_name.strip().lower():
                    return fpath
        except Exception:
            continue

    return None
