import os
import re
import sys
import json
import argparse
import tempfile
import shutil
import subprocess
import datetime
from typing import List, Dict, Set, Optional, Tuple
from urllib.parse import urlparse
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# --- CONFIGURATION ---
SPREADSHEET_ID = "10drtwlduGC1-5V1yNYFNtvBTfHM4Klre25Kzfbmirxw"

DEFAULT_SHEET_NAME = "Current Site Pages & New Site Pages"

DEFAULT_START_ROW = 26

DEFAULT_BATCH_SIZE = 10

# State persistence file
PROGRESS_FILE = "progress.json"

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

# --- PROGRESS & STATE MANAGEMENT ---

def load_progress() -> Dict:
    """Loads the progress state from PROGRESS_FILE if it exists."""
    if os.path.isfile(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Could not read {PROGRESS_FILE}: {e}")
    return {}

def update_progress(**kwargs) -> Dict:
    """Updates and saves state in PROGRESS_FILE."""
    state = load_progress()
    for k, v in kwargs.items():
        if v is not None:
            state[k] = v
    state["last_updated"] = datetime.datetime.now().isoformat()
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"⚠️ Could not save progress to {PROGRESS_FILE}: {e}")
    return state

def reset_progress():
    """Resets the state file."""
    if os.path.isfile(PROGRESS_FILE):
        try:
            os.remove(PROGRESS_FILE)
        except Exception:
            pass

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
        if creds and creds.expired and creds.refresh_token:
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

        # 3. Substring match (e.g. "Current Site Pages & New Site Pages")
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
    if not url:
        return None
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

def get_doc_links_and_md(docs_service, drive_service, doc_id: str, live_url: str = "", old_url: str = ""):
    """
    1. Reads inline links directly from Google Docs document structure.
       - Filters out phantom whitespace links (where anchor is empty or pure whitespace).
       - Filters out self-referencing canonical page links.
       - Filters out metadata links (where anchor is identical to raw URL, e.g. 'Take from here: https://...').
       - Deduplicates links.
    2. Downloads plain text representation for Antigravity diffing.
    """
    doc = docs_service.documents().get(documentId=doc_id).execute()
    links = []
    seen = set()

    # Self-referencing paths to ignore (canonical page URLs)
    self_paths = {p for p in [normalize_url_path(live_url), normalize_url_path(old_url)] if p}

    for element in doc.get("body", {}).get("content", []):
        for paragraph_element in element.get("paragraph", {}).get("elements", []):
            text_run = paragraph_element.get("textRun", {})
            link = text_run.get("textStyle", {}).get("link", {})
            if link and "url" in link:
                raw_content = text_run.get("content", "")
                anchor = raw_content.strip()
                url = link["url"].strip()

                # 1. Ignore phantom links attached to whitespace or formatting gaps
                if not anchor:
                    continue

                # 2. Ignore self-referencing links pointing to the current page itself
                url_path = normalize_url_path(url)
                if url_path in self_paths:
                    continue

                # 3. Ignore raw header/footer reference URLs (e.g. 'Take from here: https://...')
                if anchor == url or anchor == url.rstrip("/") or f"{anchor}/" == url:
                    continue

                # Deduplicate identical (anchor, url) pairs
                link_key = (anchor.lower(), url_path or url.lower())
                if link_key in seen:
                    continue
                seen.add(link_key)

                links.append({
                    "anchor": anchor,
                    "url": url
                })

    # Download doc as text/plain or markdown via Drive export
    exported = drive_service.files().export(
        fileId=doc_id, mimeType="text/plain"
    ).execute().decode("utf-8")

    return links, exported

# --- JEKYLL CONTENT MATCHING ---

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

def extract_jekyll_links(filepath: str) -> Set[str]:
    """Extracts all markdown, raw URLs, and YAML link attributes from a Jekyll file."""
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    links = set()
    # Matches markdown links: [anchor](https://... or /path)
    for m in re.findall(r"\[.*?\]\((https?://[^\s\)]+|/[^\s\)]*)\)", content):
        links.add(m.strip())

    # Matches raw URLs not enclosed in markdown brackets
    for m in re.findall(r"(?<!\()(https?://[^\s\)\"'>]+)", content):
        links.add(m.strip())

    # Matches YAML url attributes (e.g. url: "/path", link_url: /path)
    for m in re.findall(r"^\s*(?:url|button_url|link_url):\s*[\"']?([^\s\"'#]+)", content, re.MULTILINE):
        links.add(m.strip())

    return links

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

def is_link_present(doc_url: str, jekyll_links: Set[str]) -> bool:
    """Checks if a Google Doc URL is present in the extracted Jekyll links (supports absolute & relative paths)."""
    if doc_url in jekyll_links or doc_url.rstrip("/") in jekyll_links or f"{doc_url.rstrip('/')}/" in jekyll_links:
        return True

    # Check normalized paths (e.g. doc URL /native-app-development-services vs Jekyll /native-app-development-services)
    doc_path = normalize_url_path(doc_url)
    if not doc_path:
        return False

    for j_link in jekyll_links:
        if normalize_url_path(j_link) == doc_path:
            return True

    return False

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

# --- AGENT AUTOMATION WITH SESSION PERSISTENCE ---

def assign_agent_task(
    local_file: str,
    doc_path: str,
    missing_links: List[Dict[str, str]],
    jekyll_dir: str = JEKYLL_PROJECT_DIR,
    mode: str = "auto",
    conversation_id: Optional[str] = None
) -> Tuple[bool, Optional[str]]:
    """
    Assigns task to Antigravity CLI (agy) to automatically insert missing links into the Jekyll file.
    Reuses the existing conversation session for much faster execution and shared context.
    Returns (success, active_conversation_id).
    """
    missing_list_str = "\n".join([f"- Anchor: \"{m['anchor']}\" -> Target URL: {m['url']}" for m in missing_links])
    prompt = (
        f"@{local_file}\n"
        f"In project md file, add the missing links which are mentioned in the doc md file - @{doc_path}\n\n"
        f"Missing links detected:\n{missing_list_str}\n\n"
        f"Instructions:\n"
        f"1. In the target file ({local_file}), locate where the anchors or relevant sections are and insert the missing links (using root-relative slugs like /service-name for internal links).\n"
        f"2. If any of the slug links don't exist as permalink in this project, find other appropriate link which do exist.\n"
        f"3. If you couldn't find that too, list them.\n"
        f"4. Apply the edits directly to {local_file}.\n"
        f"5. IMPORTANT OUTPUT RULE: When finished, respond ONLY with the exact single line below and NOTHING else (no summaries, no explanations, no diffs):\n"
        f"All missing links have been added to {local_file}."
    )

    agy_bin = shutil.which("agy")
    if not agy_bin:
        print("❌ 'agy' (Antigravity CLI) executable not found in PATH.")
        print("Run manually with prompt:\n")
        print(prompt)
        return False, conversation_id

    cmd = [agy_bin]
    if mode == "auto":
        cmd.extend(["-p", prompt, "--mode", "accept-edits", "--dangerously-skip-permissions"])
    else:
        cmd.extend(["-i", prompt])

    # Session persistence: continue previous conversation if available
    active_id = conversation_id or get_latest_agy_conversation_id()
    if active_id:
        cmd.extend(["--conversation", active_id])
        session_info = f"session '{active_id}'"
    else:
        session_info = "new session"

    if os.path.isdir(jekyll_dir):
        cmd.extend(["--add-dir", jekyll_dir])

    print(f"🤖 Assigning task to agent ({session_info})...")

    detected_conv_id = active_id
    success = False

    try:
        if mode == "auto":
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace"
            )
            full_output = []
            for line in process.stdout:
                full_output.append(line)
                m = re.search(r"--conversation[= ]([a-f0-9-]{36})", line)
                if m:
                    detected_conv_id = m.group(1)
            process.wait()
            success = (process.returncode == 0)

            if not detected_conv_id:
                detected_conv_id = get_latest_agy_conversation_id()

            if success:
                print(f"All missing links have been added to {local_file}.\n")
            else:
                print(f"⚠️ Agent process completed with returncode non-zero.")
                tail = "".join(full_output[-10:])
                if tail:
                    print(tail)
        else:
            result = subprocess.run(cmd)
            success = (result.returncode == 0)
            detected_conv_id = get_latest_agy_conversation_id() or active_id
            if success:
                print(f"All missing links have been added to {local_file}.\n")

    except KeyboardInterrupt:
        print(f"\n⏸️ Agent interaction cancelled by user.")
        detected_conv_id = get_latest_agy_conversation_id() or active_id
        return False, detected_conv_id
    except Exception as e:
        print(f"⚠️ Failed to launch Antigravity agent: {e}")
        return False, active_id

    return success, detected_conv_id

# --- MAIN BATCH PROCESSOR ---

def process_batch(
    sheet_name: str = DEFAULT_SHEET_NAME,
    start_row: Optional[int] = None,
    batch_size: int = DEFAULT_BATCH_SIZE,
    jekyll_dir: str = JEKYLL_PROJECT_DIR,
    folders: Optional[List[str]] = None,
    agent_mode: str = "auto",
    conversation_id: Optional[str] = None,
    skip_rows: Optional[List[int]] = None
):
    # Load saved state
    state = load_progress()
    active_skip_rows = set(skip_rows if skip_rows is not None else ROWS_TO_SKIP)

    # Determine starting row: explicit CLI arg > progress.json next_row > DEFAULT_START_ROW
    if start_row is None:
        saved_row = state.get("next_row")
        if saved_row:
            start_row = saved_row
            print(f"📌 Resuming from Row {start_row} (saved in {PROGRESS_FILE}).")
        else:
            start_row = DEFAULT_START_ROW
            print(f"📌 Starting from default Row {start_row}.")
    else:
        print(f"📌 Starting from specified Row {start_row}.")

    # Determine agent session to reuse
    active_conv_id = conversation_id or state.get("conversation_id")
    if active_conv_id:
        print(f"🔗 Reusing agent session: {active_conv_id}")

    sheets, drive, docs = get_services()
    resolved_sheet = get_resolved_sheet_name(sheets, SPREADSHEET_ID, sheet_name)

    # Fetch rows starting from row 1 (A1:F)
    sheet_range = f"'{resolved_sheet}'!A1:F"
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=sheet_range
    ).execute()
    rows = result.get("values", [])

    total_rows = len(rows)
    if start_row > total_rows:
        print(f"⚠️ Start row {start_row} is greater than total rows ({total_rows}) in sheet '{resolved_sheet}'.")
        return

    current_start = start_row
    current_row_num = start_row

    try:
        while current_start <= total_rows:
            if batch_size and batch_size > 0:
                end_row = min(current_start + batch_size - 1, total_rows)
            else:
                end_row = total_rows

            print("---\n" * 4)
            print(f"\n🚀 Processing '{resolved_sheet}' - Rows {current_start} to {end_row} (Batch: {end_row - current_start + 1}, Total: {total_rows})...\n")

            for row_idx in range(current_start - 1, end_row):
                current_row_num = row_idx + 1
                row = rows[row_idx]

                if current_row_num in active_skip_rows:
                    print(f"⏩ [Row {current_row_num}] Skipped as configured in ROWS_TO_SKIP.")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    continue

                # Column layout for "Current Site Pages & New Site Pages":
                # Col A (0): Category (e.g. Services & Other Pages)
                # Col B (1): Old Site Pages (URL)
                # Col C (2): Content Link (Google Doc URL)
                # Col D (3): Status (Done)
                # Col E (4): New Site Pages (Staging/Live URL)
                # Col F (5): Jay Status
                category = row[0] if len(row) > 0 else ""
                old_url = row[1] if len(row) > 1 else ""
                doc_col = row[2] if len(row) > 2 else ""
                status = row[3] if len(row) > 3 else ""
                new_url = row[4] if len(row) > 4 else ""
                jay_status = row[5] if len(row) > 5 else ""

                # Primary live URL is from New Site Pages (Col E), fallback to Old Site Pages (Col B)
                live_link = new_url if (new_url and new_url.startswith("http")) else old_url

                # Derive human-readable page name from slug
                if live_link.startswith("http"):
                    slug_part = urlparse(live_link).path.strip("/").split("/")[-1]
                    page_name = slug_part.replace("-", " ").title() if slug_part else f"Row {current_row_num}"
                else:
                    page_name = old_url if old_url else f"Row {current_row_num}"

                # Extract Google Doc ID from Column C (or scan any cell in row as fallback)
                doc_id = extract_doc_id(doc_col)
                if not doc_id:
                    for cell in row:
                        doc_id = extract_doc_id(cell)
                        if doc_id:
                            break

                if not doc_id:
                    print(f"⏩ [Row {current_row_num}] {page_name}: No valid Google Doc found in Column C. Skipping.")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    continue

                print(f"--- Checking [Row {current_row_num}] {page_name} ---")

                # 1. Extract Links from Google Doc
                doc_links, doc_md_text = get_doc_links_and_md(docs, drive, doc_id, live_url=live_link, old_url=old_url)
                if not doc_links:
                    print(f"ℹ️ No links found in Google Doc for '{page_name}'.")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    continue

                # 2. Find Local Jekyll File across configured content folders
                local_file = find_matching_jekyll_file(page_name, live_link, jekyll_dir=jekyll_dir, folders=folders)
                if not local_file:
                    print(f"⚠️ Could not find local .md file for '{page_name}' (Slug: {urlparse(live_link).path}).")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    continue

                # 3. Check for Missing Links
                jekyll_urls = extract_jekyll_links(local_file)
                missing_links = [l for l in doc_links if not is_link_present(l["url"], jekyll_urls)]

                if not missing_links:
                    print(f"✅ All {len(doc_links)} links are present in `{local_file}`.\n")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    continue

                print(f"❌ Found {len(missing_links)} missing link(s):")
                for m in missing_links:
                    print(f"   - [{m['anchor']}] -> {m['url']}")

                # 4. Prepare doc markdown in _work_docs inside the Jekyll project
                work_docs_dir = os.path.join(jekyll_dir, "_work_docs")
                clean_name = re.sub(r'[\/:*?"<>|]', '_', page_name).strip()
                try:
                    os.makedirs(work_docs_dir, exist_ok=True)
                    doc_path = os.path.join(work_docs_dir, f"{clean_name}.md")
                except Exception:
                    doc_path = os.path.join(tempfile.gettempdir(), f"{clean_name}.md")

                with open(doc_path, "w", encoding="utf-8") as f:
                    f.write(doc_md_text)

                # 5. Automatically assign task to Antigravity Agent
                if agent_mode == "skip":
                    print("⏩ Skipping agent task as requested.\n")
                else:
                    _, active_conv_id = assign_agent_task(
                        local_file=local_file,
                        doc_path=doc_path,
                        missing_links=missing_links,
                        jekyll_dir=jekyll_dir,
                        mode=agent_mode,
                        conversation_id=active_conv_id
                    )

                # Save progress after completing this row
                update_progress(
                    last_processed_row=current_row_num,
                    next_row=current_row_num + 1,
                    sheet_name=resolved_sheet,
                    conversation_id=active_conv_id
                )

            # Check if all rows in sheet have been processed
            if end_row >= total_rows:
                print(f"\n🎉 All {total_rows} rows in '{resolved_sheet}' have been processed!")
                break

            # Prompt to continue to next batch (default: Y)
            next_batch_start = end_row + 1
            next_batch_end = min(next_batch_start + batch_size - 1, total_rows) if batch_size and batch_size > 0 else total_rows
            action = input(f"Batch completed up to Row {end_row}. Continue to next batch (Rows {next_batch_start}-{next_batch_end})? [Y/n]: ").strip().lower()
            if action not in ["", "y", "yes"]:
                print(f"💾 Progress saved. Next run will resume from Row {next_batch_start}.")
                break

            current_start = next_batch_start

    except KeyboardInterrupt:
        print(f"\n\n⏸️ Interrupted by user. Progress saved at Row {current_row_num}.")
        update_progress(
            last_processed_row=current_row_num - 1,
            next_row=current_row_num,
            sheet_name=resolved_sheet,
            conversation_id=active_conv_id
        )
        print(f"Next run will resume from Row {current_row_num}.")
        sys.exit(0)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Batch link verifier")
    parser.add_argument("--sheet", type=str, default=DEFAULT_SHEET_NAME, help=f"Sheet/tab name to process (default: '{DEFAULT_SHEET_NAME}')")
    parser.add_argument("--start", type=int, default=None, help=f"Starting row number (default: resume from {PROGRESS_FILE} or {DEFAULT_START_ROW})")
    parser.add_argument("--batch", type=int, default=DEFAULT_BATCH_SIZE, help="Batch size (number of rows to process; 0 for all remaining)")
    parser.add_argument("--jekyll-dir", type=str, default=JEKYLL_PROJECT_DIR, help="Absolute path to Jekyll project directory")
    parser.add_argument("--folders", nargs="+", default=CONTENT_FOLDERS, help="Content folders to scan (e.g. _services _technologies)")
    parser.add_argument(
        "--agent-mode",
        choices=["auto", "interactive", "skip"],
        default="auto",
        help="Agent execution mode: 'auto' (direct automated edits, default), 'interactive', or 'skip'"
    )
    parser.add_argument("--conversation", type=str, default=None, help="Explicit Antigravity conversation ID to resume")
    parser.add_argument("--new-session", action="store_true", help="Start a new agent conversation session instead of resuming")
    parser.add_argument("--skip-rows", nargs="+", type=int, default=ROWS_TO_SKIP, help=f"Row numbers to skip (default: {ROWS_TO_SKIP})")
    parser.add_argument("--reset", action="store_true", help=f"Reset saved progress in {PROGRESS_FILE} and start from {DEFAULT_START_ROW}")
    args = parser.parse_args()

    if args.reset:
        reset_progress()
        print(f"🔄 Progress state reset. Starting from Row {DEFAULT_START_ROW}.")

    process_batch(
        sheet_name=args.sheet,
        start_row=args.start,
        batch_size=args.batch,
        jekyll_dir=args.jekyll_dir,
        folders=args.folders,
        agent_mode=args.agent_mode,
        conversation_id=args.conversation if not args.new_session else "",
        skip_rows=args.skip_rows
    )