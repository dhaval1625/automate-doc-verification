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

from common import (
    SPREADSHEET_ID,
    DEFAULT_SHEET_NAME,
    DEFAULT_START_ROW,
    LAST_ROW_TO_CHECK,
    DEFAULT_BATCH_SIZE,
    JEKYLL_PROJECT_DIR,
    CONTENT_FOLDERS,
    SCOPES,
    ROWS_TO_SKIP,
    get_latest_agy_conversation_id,
    get_services,
    get_resolved_sheet_name,
    extract_doc_id,
    normalize_url_path,
    get_target_directories,
    find_matching_jekyll_file,
)

# State persistence file
PROGRESS_FILE = "progress.json"

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
            actionable_count = 0
            empty_docs_skipped = 0
            batch_target = batch_size if (batch_size and batch_size > 0) else None

            target_desc = f"{batch_target} actionable pages" if batch_target else "all remaining rows"
            print("---\n" * 4)
            print(f"\n🚀 Processing '{resolved_sheet}' starting from Row {current_start} (Target: {target_desc}, Total Sheet Rows: {total_rows})...\n")

            row_idx = current_start - 1
            while row_idx < total_rows:
                current_row_num = row_idx + 1
                row = rows[row_idx]

                if current_row_num in active_skip_rows:
                    if empty_docs_skipped > 0:
                        print(f"⏩ Skipped {empty_docs_skipped} row(s) without Google Doc links.")
                        empty_docs_skipped = 0
                    print(f"⏩ [Row {current_row_num}] Skipped as configured in ROWS_TO_SKIP.")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    continue

                # Column layout for "Current Site Pages & New Site Pages":
                # Col C (2): Content Link (Google Doc URL)
                old_url = row[1] if len(row) > 1 else ""
                doc_col = row[2] if len(row) > 2 else ""
                new_url = row[4] if len(row) > 4 else ""

                # Extract Google Doc ID from Column C (or scan any cell in row as fallback)
                doc_id = extract_doc_id(doc_col)
                if not doc_id:
                    for cell in row:
                        doc_id = extract_doc_id(cell)
                        if doc_id:
                            break

                # If no doc link is found, skip quietly in memory
                if not doc_id:
                    empty_docs_skipped += 1
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    continue

                # Actionable row found
                if empty_docs_skipped > 0:
                    print(f"⏩ Skipped {empty_docs_skipped} row(s) without Google Doc links.")
                    empty_docs_skipped = 0

                actionable_count += 1
                live_link = new_url if (new_url and new_url.startswith("http")) else old_url

                # Derive human-readable page name from slug
                if live_link.startswith("http"):
                    slug_part = urlparse(live_link).path.strip("/").split("/")[-1]
                    page_name = slug_part.replace("-", " ").title() if slug_part else f"Row {current_row_num}"
                else:
                    page_name = old_url if old_url else f"Row {current_row_num}"

                batch_info = f"[{actionable_count}/{batch_target}] " if batch_target else ""
                print(f"--- Checking {batch_info}[Row {current_row_num}] {page_name} ---")

                # 1. Extract Links from Google Doc
                doc_links, doc_md_text = get_doc_links_and_md(docs, drive, doc_id, live_url=live_link, old_url=old_url)
                if not doc_links:
                    print(f"ℹ️ No links found in Google Doc for '{page_name}'.\n")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    if batch_target and actionable_count >= batch_target:
                        break
                    continue

                # 2. Find Local Jekyll File across configured content folders
                local_file = find_matching_jekyll_file(page_name, live_link, jekyll_dir=jekyll_dir, folders=folders)
                if not local_file:
                    print(f"⚠️ Could not find local .md file for '{page_name}' (Slug: {urlparse(live_link).path}).\n")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    if batch_target and actionable_count >= batch_target:
                        break
                    continue

                # 3. Check for Missing Links
                jekyll_urls = extract_jekyll_links(local_file)
                missing_links = [l for l in doc_links if not is_link_present(l["url"], jekyll_urls)]

                if not missing_links:
                    print(f"✅ All {len(doc_links)} links are present in `{local_file}`.\n")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    if batch_target and actionable_count >= batch_target:
                        break
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

                row_idx += 1
                if batch_target and actionable_count >= batch_target:
                    break

            if empty_docs_skipped > 0:
                print(f"⏩ Skipped {empty_docs_skipped} row(s) without Google Doc links.")

            # Check if all rows in sheet have been processed
            if row_idx >= total_rows:
                print(f"\n🎉 All {total_rows} rows in '{resolved_sheet}' have been processed!")
                break

            # Prompt to continue to next batch (default: Y)
            next_batch_start = current_row_num + 1
            action = input(f"Batch completed ({actionable_count} pages processed, reached Row {current_row_num}). Continue to next batch? [Y/n]: ").strip().lower()
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