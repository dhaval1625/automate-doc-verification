import os
import re
import sys
import json
import argparse
import datetime
import subprocess
import shutil
import tempfile
from typing import List, Dict, Any, Optional, Tuple
from difflib import SequenceMatcher
from urllib.parse import urlparse
import yaml

from common import (
    SPREADSHEET_ID,
    DEFAULT_SHEET_NAME,
    DEFAULT_START_ROW,
    LAST_ROW_TO_CHECK,
    DEFAULT_BATCH_SIZE,
    JEKYLL_PROJECT_DIR,
    CONTENT_FOLDERS,
    ROWS_TO_SKIP,
    get_latest_agy_conversation_id,
    get_services,
    get_resolved_sheet_name,
    extract_doc_id,
    normalize_url_path,
    find_matching_jekyll_file,
)

# State persistence file for content verification
PROGRESS_FILE = "content_progress.json"

# Minimum similarity ratio to consider two text blocks a match
DEFAULT_SIMILARITY_THRESHOLD = 0.88

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

# --- TEXT NORMALIZATION & SIMILARITY ---

def normalize_text(text: str) -> str:
    """
    Normalizes text for robust comparison:
    - Strips markdown link syntax [anchor](url) -> anchor
    - Replaces typographic/curly quotes and apostrophes
    - Strips markdown headers, bold, bullet characters, backslash escapes
    - Replaces unicode quotes and non-breaking spaces
    - Preserves C# / F# while stripping markdown headings and formatting
    - Collapses whitespace into single space and converts to lowercase
    """
    if not text:
        return ""
    # Strip markdown link target, keep anchor text: [anchor](url) -> anchor
    text = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", str(text))
    # Replace curly quotes and unicode punctuation
    text = text.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    text = text.replace("\u00a0", " ").replace("\ufeff", "")
    # Remove markdown headings but preserve C# and F#
    text = re.sub(r"(?<![cfCF])#+", " ", text)
    # Remove markdown formatting characters
    text = re.sub(r"[\*_`\\\->|]", " ", text)
    # Collapse whitespace and lowercase
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text

def clean_item_title(title: str) -> str:
    """Strips leading numbering/bullets like '1. ', '1) ', '* '."""
    if not title:
        return ""
    t = str(title).strip()
    t = re.sub(r"^\d+[\.\)]\s*", "", t)
    t = re.sub(r"^[\*\-\s]+", "", t)
    return t.strip()

def match_titles(t1: str, t2: str) -> bool:
    """Matches two titles ignoring leading numbers, case, and punctuation."""
    c1 = normalize_text(clean_item_title(t1))
    c2 = normalize_text(clean_item_title(t2))
    if not c1 or not c2:
        return False
    if c1 == c2:
        return True
    if len(c1) > 4 and (c1 in c2 or c2 in c1):
        return True
    return text_similarity(c1, c2) >= 0.85

def text_similarity(a: str, b: str) -> float:
    """Returns SequenceMatcher ratio between two normalized strings."""
    na = normalize_text(a)
    nb = normalize_text(b)
    if not na and not nb:
        return 1.0
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()

def is_text_contained(needle: str, haystack: str, threshold: float = DEFAULT_SIMILARITY_THRESHOLD) -> bool:
    """Checks if normalized needle is present or highly similar within normalized haystack."""
    nn = normalize_text(needle)
    nh = normalize_text(haystack)
    if not nn:
        return True
    if not nh:
        return False
    if nn in nh:
        return True
    words = nn.split()
    if len(words) >= 4:
        # Check sliding window of word n-grams
        ngram_size = min(len(words), 5)
        for i in range(len(words) - ngram_size + 1):
            window = " ".join(words[i:i + ngram_size])
            if window in nh:
                return True
    return text_similarity(nn, nh) >= threshold

# --- GOOGLE DOCS API & PLAIN TEXT PARSER ---

MAJOR_SECTIONS = [
    "project overview",
    "project objectives",
    "business challenges",
    "our solutions",
    "key benefits",
    "key features",
    "technology",
    "business results",
    "frequently asked questions"
]

def parse_docs_api(doc: Dict[str, Any]) -> Tuple[Dict[str, Any], str]:
    """
    Parses Google Docs API response structure into complete data:
    - Extracts title, banner_description, industry, services, country
    - Extracts all major sections, paragraphs, bullet points, challenges, features, tech rows, results, and FAQs
    - Converts document into high-fidelity markdown with headings, bullets, tables, and links
    """
    data: Dict[str, Any] = {
        "title": "",
        "banner_description": "",
        "industry": "",
        "services": [],
        "country": "",
        "sections": []
    }

    md_lines: List[str] = []
    body = doc.get("body", {}).get("content", [])
    current_sec: Optional[Dict[str, Any]] = None
    collecting_services = False
    expecting_banner_desc = False

    for el in body:
        if "paragraph" in el:
            p = el["paragraph"]
            style = p.get("paragraphStyle", {}).get("namedStyleType", "")
            bullet = p.get("bullet")

            # Extract formatted text with links and bold marks
            p_text = ""
            is_all_bold = True
            elements = p.get("elements", [])
            bold_text = ""
            non_bold_text = ""
            seen_non_bold = False

            for tr in elements:
                t_run = tr.get("textRun", {})
                content = t_run.get("content", "")
                if not content:
                    continue
                clean_content = content.replace("\n", "")
                link = t_run.get("textStyle", {}).get("link", {}).get("url")
                bold = t_run.get("textStyle", {}).get("bold")
                if clean_content.strip() and not bold:
                    is_all_bold = False

                if bold and not seen_non_bold:
                    bold_text += clean_content
                else:
                    if clean_content.strip():
                        seen_non_bold = True
                    non_bold_text += clean_content

                if link:
                    clean_content = f"[{clean_content}]({link})"
                if bold and clean_content.strip():
                    clean_content = f"**{clean_content}**"
                p_text += clean_content

            p_text = p_text.strip()
            if not p_text:
                continue

            clean_raw_text = re.sub(r"[\*#_]", "", p_text).strip()
            norm_p = clean_raw_text.replace("\x0b", " ").lower().strip(": -")
            norm_p = re.sub(r"\s+", " ", norm_p)

            # Markdown line representation for export
            if style == "HEADING_1":
                md_lines.append(f"# {p_text}")
            elif style == "HEADING_2":
                md_lines.append(f"## {p_text}")
            elif style == "HEADING_3":
                md_lines.append(f"### {p_text}")
            elif bullet:
                md_lines.append(f"* {p_text}")
            else:
                md_lines.append(p_text)

            # 1. Project Title
            if norm_p.startswith("project name"):
                data["title"] = re.sub(r"[\*#_]", "", p_text.split(":", 1)[-1]).strip().strip("- ")
                continue
            elif not data["title"] and (style in ["HEADING_1", "TITLE"]) and not any(norm_p == ms for ms in MAJOR_SECTIONS):
                data["title"] = clean_raw_text
                continue

            # 2. Short / Banner Description
            if norm_p in ["short description", "banner description"]:
                expecting_banner_desc = True
                continue
            elif expecting_banner_desc:
                data["banner_description"] = clean_raw_text
                expecting_banner_desc = False
                continue
            elif norm_p.startswith("short description:") or norm_p.startswith("banner description:"):
                data["banner_description"] = re.sub(r"[\*#_]", "", p_text.split(":", 1)[-1]).strip()
                continue

            # 3. Industry
            if norm_p.startswith("industry:"):
                data["industry"] = re.sub(r"[\*#_]", "", p_text.split(":", 1)[-1]).strip()
                continue

            # 4. Services
            if norm_p == "services" or norm_p.startswith("services:"):
                collecting_services = True
                continue
            elif collecting_services:
                if norm_p.startswith("country:") or any(norm_p == ms for ms in MAJOR_SECTIONS) or any(norm_p.startswith(f) for f in ["industry:", "project name:", "short description:"]):
                    collecting_services = False
                else:
                    data["services"].append(re.sub(r"[\*#_]", "", p_text.lstrip("*- ")).strip())
                    continue

            # 5. Country
            if norm_p.startswith("country:"):
                data["country"] = re.sub(r"[\*#_]", "", p_text.split(":", 1)[-1]).strip()
                continue

            # 6. Major Sections
            matched_major = None
            for ms in MAJOR_SECTIONS:
                if norm_p == ms or (style in ["HEADING_1", "HEADING_2"] and ms in norm_p):
                    matched_major = ms
                    break

            if matched_major:
                clean_title = clean_raw_text
                current_sec = {
                    "title": clean_title,
                    "items": [],
                    "paragraphs": [],
                    "table_rows": [],
                    "faqs": []
                }
                data["sections"].append(current_sec)
                continue

            if not current_sec:
                continue

            sec_name = current_sec["title"].lower()

            # FAQs
            if "faq" in sec_name or "frequently asked" in sec_name:
                m_q = re.match(r"^\d+[\)\.]\s*(.+)$", clean_raw_text)
                if m_q or clean_raw_text.endswith("?"):
                    q_text = m_q.group(1).strip() if m_q else clean_raw_text
                    current_sec["faqs"].append({"question": q_text, "answer": ""})
                elif current_sec["faqs"]:
                    if current_sec["faqs"][-1]["answer"]:
                        current_sec["faqs"][-1]["answer"] += " " + clean_raw_text
                    else:
                        current_sec["faqs"][-1]["answer"] = clean_raw_text
                continue

            # Business Results
            if "business results" in sec_name:
                m_res = re.search(r"(\d+%\+?)\s*(.+)", clean_raw_text)
                if m_res and ("HEADING" in style or len(clean_raw_text) < 70):
                    current_sec["items"].append({
                        "value": m_res.group(1).strip(),
                        "text": m_res.group(2).strip(),
                        "subtext": ""
                    })
                elif current_sec["items"] and isinstance(current_sec["items"][-1], dict) and "subtext" in current_sec["items"][-1]:
                    if current_sec["items"][-1]["subtext"]:
                        current_sec["items"][-1]["subtext"] += " " + clean_raw_text
                    else:
                        current_sec["items"][-1]["subtext"] = clean_raw_text
                continue

            # Key Features or Challenges
            if "feature" in sec_name or "challenge" in sec_name:
                # Check 1: Inline bold title + description
                clean_b = bold_text.replace("\x0b", " ").strip()
                clean_nb = non_bold_text.replace("\x0b", " ").strip()
                if clean_b and clean_nb and len(clean_b.split()) <= 12:
                    current_sec["items"].append({
                        "title": clean_b,
                        "description": clean_nb
                    })
                    continue
                # Check 2: Soft break (\x0b) splitting title and description
                if "\x0b" in clean_raw_text:
                    parts = [pt.strip() for pt in clean_raw_text.split("\x0b") if pt.strip()]
                    if len(parts) >= 2 and len(parts[0].split()) <= 12:
                        current_sec["items"].append({
                            "title": parts[0],
                            "description": " ".join(parts[1:])
                        })
                        continue
                m_num = re.match(r"^\d+[\.\)]\s*(.+)$", clean_raw_text)
                is_bold_title = is_all_bold or (style in ["HEADING_2", "HEADING_3"])
                is_item_title = bool(m_num) or (is_bold_title and len(clean_raw_text.split()) <= 8)
                if is_item_title:
                    item_title = m_num.group(1).strip() if m_num else clean_raw_text
                    current_sec["items"].append({
                        "title": item_title,
                        "description": ""
                    })
                elif current_sec["items"] and isinstance(current_sec["items"][-1], dict) and "description" in current_sec["items"][-1]:
                    if current_sec["items"][-1]["description"]:
                        current_sec["items"][-1]["description"] += " " + clean_raw_text
                    else:
                        current_sec["items"][-1]["description"] = clean_raw_text
                else:
                    current_sec["paragraphs"].append(clean_raw_text)
                continue

            # Bullet items
            if bullet or p_text.startswith("*") or p_text.startswith("-"):
                current_sec["items"].append(re.sub(r"[\*#_]", "", p_text.lstrip("*- ")).strip())
                continue

            # Standard Paragraph
            current_sec["paragraphs"].append(clean_raw_text)

        elif "table" in el:
            table = el["table"]
            for row in table.get("tableRows", []):
                cells = []
                for cell in row.get("tableCells", []):
                    c_txt = ""
                    for cell_el in cell.get("content", []):
                        if "paragraph" in cell_el:
                            for tr in cell_el["paragraph"].get("elements", []):
                                c_txt += tr.get("textRun", {}).get("content", "")
                    cells.append(c_txt.strip().replace("\n", " "))
                md_lines.append("| " + " | ".join(cells) + " |")
                if current_sec and len(cells) >= 2 and cells[0].lower() not in ["category", "label"]:
                    current_sec["table_rows"].append({"label": cells[0], "value": cells[1]})

    return data, "\n\n".join(md_lines)

# --- GOOGLE DOC CONTENT PARSER ---

def parse_doc_content(doc_text: str) -> Dict[str, Any]:
    """
    Parses Google Doc markdown / plain text into structured sections, metadata, and items.
    Extracts:
    - title
    - banner_description
    - industry, services, country
    - ordered sections (Overview, Objectives, Challenges, Solutions, Features, Technology, Results, FAQs)
    """
    lines = doc_text.splitlines()
    data: Dict[str, Any] = {
        "title": "",
        "banner_description": "",
        "industry": "",
        "services": [],
        "country": "",
        "sections": []
    }

    current_section: Optional[Dict[str, Any]] = None
    collecting_services = False
    expecting_banner_desc = False

    for idx, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue

        # Normalize common markdown escapes in doc lines (e.g. 1\. -> 1.)
        line = re.sub(r"\\([.\-\(\)])", r"\1", line)
        norm_line = normalize_text(line)

        # 1. Project Title: Project Name: ...
        m_title = re.search(r"(?:project name|title)\s*:\s*(.*)", line, re.IGNORECASE)
        if m_title and not data["title"]:
            raw_title = re.sub(r"[\*#_]", "", m_title.group(1)).strip().strip("- ")
            data["title"] = raw_title
            continue
        elif line.startswith("# ") and not data["title"] and not m_title:
            data["title"] = re.sub(r"[\*#_]", "", line).strip()
            continue

        # 2. Short / Banner Description
        m_desc = re.search(r"(?:short description|banner description)\s*:\s*(.*)", line, re.IGNORECASE)
        if m_desc and not data["banner_description"]:
            data["banner_description"] = re.sub(r"[\*#_]", "", m_desc.group(1)).strip()
            continue
        elif re.search(r"^(?:#*\s*\*{0,2})?(?:short description|banner description)(?:\*{0,2})?$", line, re.IGNORECASE):
            expecting_banner_desc = True
            continue
        elif expecting_banner_desc:
            data["banner_description"] = re.sub(r"[\*#_]", "", line).strip()
            expecting_banner_desc = False
            continue

        # 3. Industry
        m_ind = re.search(r"industry\s*:\s*(.*)", line, re.IGNORECASE)
        if m_ind and not data["industry"]:
            data["industry"] = re.sub(r"[\*#_]", "", m_ind.group(1)).strip()
            continue

        # 4. Country
        m_country = re.search(r"country\s*:\s*(.*)", line, re.IGNORECASE)
        if m_country and not data["country"]:
            data["country"] = re.sub(r"[\*#_]", "", m_country.group(1)).strip()
            continue

        # 5. Services header
        if re.search(r"^Services\s*:", line, re.IGNORECASE):
            collecting_services = True
            continue

        if collecting_services:
            if line.startswith("*") or line.startswith("-"):
                item = re.sub(r"^[\*\-\s]+", "", line).strip()
                if item:
                    data["services"].append(item)
                continue
            else:
                collecting_services = False

        # 6. Check for Major Section Headings
        # Check if line is a sub-metric in Business Results (e.g. ### **35% Faster...**)
        clean_metric_cand = re.sub(r"^[#\*\s]+|[#\*\s]+$", "", line)
        m_metric = re.match(r"^(\d+%\+?)\s*(.+)$", clean_metric_cand)
        if current_section and "business results" in current_section["title"].lower() and m_metric:
            val = m_metric.group(1).strip()
            txt = re.sub(r"[\*#_]", "", m_metric.group(2)).strip()
            current_section["items"].append({
                "value": val,
                "text": txt,
                "subtext": ""
            })
            continue

        # Headings for Major Sections
        clean_head_cand = re.sub(r"[\*#_]", "", line).strip()
        norm_head = clean_head_cand.lower()
        matched_ms = None
        for ms in MAJOR_SECTIONS:
            if norm_head == ms or (norm_head.startswith(ms) and len(norm_head.split()) <= 4):
                matched_ms = ms
                break

        if matched_ms:
            current_section = {
                "title": clean_head_cand,
                "raw_heading": line,
                "items": [],
                "paragraphs": [],
                "table_rows": [],
                "faqs": []
            }
            data["sections"].append(current_section)
            continue

        if not current_section:
            # Intro text before any explicit heading
            if len(line) > 20 and not data["banner_description"]:
                data["banner_description"] = line
            continue

        sec_name = current_section["title"].lower()

        # 7. FAQs parsing: **1) Question?** \n Answer
        m_faq = re.match(r"^\*{0,2}\d+[\)\.]\s*(.+?)\*{0,2}$", line)
        if "faq" in sec_name or "frequently asked questions" in sec_name:
            if m_faq:
                q_text = m_faq.group(1).strip()
                current_section["faqs"].append({
                    "question": q_text,
                    "answer": ""
                })
            elif current_section["faqs"]:
                if current_section["faqs"][-1]["answer"]:
                    current_section["faqs"][-1]["answer"] += " " + line
                else:
                    current_section["faqs"][-1]["answer"] = line
            continue

        # 8. Technology table rows: | Frontend | ASP.NET ... |
        if "technology" in sec_name:
            if line.startswith("|"):
                cells = [c.strip() for c in line.split("|") if c.strip()]
                if cells and not all(set(c) <= set("- ") for c in cells):
                    if len(cells) >= 2 and cells[0].lower() not in ["category", "label"]:
                        current_section["table_rows"].append({
                            "label": cells[0],
                            "value": cells[1]
                        })
                continue

        # 9. Business Results metrics: ### **35% Faster Store Operations** \n text
        if "business results" in sec_name:
            m_metric = re.match(r"^\*{0,3}(\d+%\+?)\s*(.+?)\*{0,3}$", line)
            if m_metric:
                val = m_metric.group(1).strip()
                txt = m_metric.group(2).strip()
                current_section["items"].append({
                    "value": val,
                    "text": txt,
                    "subtext": ""
                })
                continue
            elif current_section["items"]:
                last_res = current_section["items"][-1]
                if isinstance(last_res, dict) and "subtext" in last_res:
                    if last_res["subtext"]:
                        last_res["subtext"] += " " + line
                    else:
                        last_res["subtext"] = line
                    continue

        # 10. Numbered or Bolded challenges/features: **1. Title** \n description or **Title **Description
        if "challenge" in sec_name or "feature" in sec_name:
            m_inline = re.match(r"^\s*(?:\d+[\.\)]\s*)?\*\*(.+?)\*\*[:\s\-]*(.+)$", line)
            if m_inline:
                t_cand = re.sub(r"[\*#_]", "", m_inline.group(1)).strip()
                d_cand = re.sub(r"[\*#_]", "", m_inline.group(2)).strip()
                if len(t_cand.split()) <= 12 and d_cand:
                    current_section["items"].append({
                        "title": t_cand,
                        "description": d_cand
                    })
                    continue

            m_item_title = re.match(r"^(?:#{1,4}\s*)?\*{0,2}(?:\d+[\.\)]\s*)?([A-Z0-9][A-Za-z0-9\s&/\-]+)\*{0,2}$", line)
            if m_item_title and len(m_item_title.group(1).split()) <= 8:
                current_section["items"].append({
                    "title": m_item_title.group(1).strip(),
                    "description": ""
                })
                continue
            elif current_section["items"]:
                last_item = current_section["items"][-1]
                if isinstance(last_item, dict) and "description" in last_item:
                    if last_item["description"]:
                        last_item["description"] += " " + line
                    else:
                        last_item["description"] = line
                    continue

        # 11. Bullet items: * Bullet point
        if line.startswith("*") or line.startswith("-"):
            bullet = re.sub(r"^[\*\-\s]+", "", line).strip()
            if bullet:
                current_section["items"].append(bullet)
            continue

        # 12. Standard Paragraph
        if len(line) > 10:
            current_section["paragraphs"].append(line)

    return data

# --- JEKYLL FRONT MATTER PARSER ---

def parse_jekyll_front_matter(filepath: str) -> Dict[str, Any]:
    """
    Parses only the YAML Front Matter (between first and second ---) of a Jekyll markdown file.
    Extracts:
    - title, banner_description, industry, services, country
    - sections list: [ {type, title, content, items, rows, points} ]
    """
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    parts = content.split("---")
    if len(parts) < 3:
        return {}

    front_matter_raw = parts[1]
    try:
        data = yaml.safe_load(front_matter_raw) or {}
    except Exception as e:
        print(f"⚠️ YAML parsing error in {filepath}: {e}")
        return {}

    return data

# --- CONTENT COMPARATOR & VERIFICATION ENGINE ---

def compare_content(
    doc_data: Dict[str, Any],
    jekyll_data: Dict[str, Any],
    threshold: float = DEFAULT_SIMILARITY_THRESHOLD
) -> Tuple[float, List[Dict[str, str]], List[str]]:
    """
    Compares Google Doc specification with Jekyll Front Matter:
    1. Metadata audit (title, banner_description, services, industry, country)
    2. Section-by-section and item-by-item strictly ordered verification.
       Enforces:
       - Exact sequence order of points
       - Sublabels matching (e.g. description under title, answer under question)
       - Detects removed, shortened, swapped, or modified descriptions
    Returns:
    - match_score: float (0.0 to 100.0)
    - discrepancies: List of detailed discrepancy objects
    - success_items: List of matched items summary
    """
    total_checks = 0
    passed_checks = 0
    discrepancies: List[Dict[str, str]] = []
    success_items: List[str] = []

    # --- 1. METADATA CHECKS ---

    # Title check
    if doc_data.get("title"):
        total_checks += 1
        j_title = str(jekyll_data.get("title", ""))
        if match_titles(doc_data["title"], j_title):
            passed_checks += 1
            success_items.append("Title matched")
        else:
            discrepancies.append({
                "type": "METADATA_MISMATCH",
                "field": "title",
                "expected": doc_data["title"],
                "found": j_title,
                "message": f"Title mismatch. Expected: '{doc_data['title']}' | Found: '{j_title}'"
            })

    # Banner Description check
    if doc_data.get("banner_description"):
        total_checks += 1
        j_desc = str(jekyll_data.get("banner_description", "") or jekyll_data.get("description", ""))
        norm_d = normalize_text(doc_data["banner_description"])
        norm_j = normalize_text(j_desc)
        sim = text_similarity(norm_d, norm_j)
        if norm_d == norm_j or sim >= 0.85:
            passed_checks += 1
            success_items.append("Banner description matched")
        else:
            discrepancies.append({
                "type": "METADATA_MISMATCH",
                "field": "banner_description",
                "expected": doc_data["banner_description"][:80] + "...",
                "found": j_desc[:80] + "...",
                "message": f"Banner description mismatch. Expected: '{doc_data['banner_description'][:80]}...' | Found: '{j_desc[:80]}...'"
            })

    # Industry check
    if doc_data.get("industry"):
        total_checks += 1
        j_ind = str(jekyll_data.get("industry", "") or jekyll_data.get("category", ""))
        norm_d = normalize_text(doc_data["industry"])
        norm_j = normalize_text(j_ind)
        if norm_d in norm_j or norm_j in norm_d or text_similarity(norm_d, norm_j) >= 0.80:
            passed_checks += 1
            success_items.append("Industry matched")
        else:
            discrepancies.append({
                "type": "METADATA_MISMATCH",
                "field": "industry",
                "expected": doc_data["industry"],
                "found": j_ind,
                "message": f"Industry mismatch. Expected: '{doc_data['industry']}' | Found: '{j_ind}'"
            })

    # Country check
    if doc_data.get("country"):
        total_checks += 1
        j_country = str(jekyll_data.get("country", ""))
        norm_d = normalize_text(doc_data["country"])
        norm_j = normalize_text(j_country)
        if norm_d in norm_j or norm_j in norm_d:
            passed_checks += 1
        else:
            discrepancies.append({
                "type": "METADATA_MISMATCH",
                "field": "country",
                "expected": doc_data["country"],
                "found": j_country,
                "message": f"Country mismatch. Expected: '{doc_data['country']}' | Found: '{j_country}'"
            })

    # Services check (ordered)
    if doc_data.get("services"):
        j_services = jekyll_data.get("services", [])
        if not isinstance(j_services, list):
            j_services = [str(j_services)]
        
        last_s_idx = -1
        matched_s_indices = set()
        for idx, serv in enumerate(doc_data["services"]):
            total_checks += 1
            matching_idx = None
            if idx < len(j_services) and match_titles(serv, j_services[idx]):
                matching_idx = idx
            else:
                for j_idx, js in enumerate(j_services):
                    if j_idx not in matched_s_indices and match_titles(serv, js):
                        matching_idx = j_idx
                        break
            if matching_idx is None:
                discrepancies.append({
                    "type": "METADATA_MISMATCH",
                    "field": "services",
                    "expected": serv,
                    "found": ", ".join(j_services),
                    "message": f"Service '{serv}' missing from Jekyll services front matter."
                })
            else:
                matched_s_indices.add(matching_idx)
                total_checks += 1
                if matching_idx < last_s_idx or matching_idx != idx:
                    discrepancies.append({
                        "type": "ORDER_MISMATCH",
                        "field": "services",
                        "item": serv,
                        "message": f"Service '{serv}' is out of order (expected position #{idx+1}, found at #{matching_idx+1})."
                    })
                else:
                    passed_checks += 1
                passed_checks += 1
                last_s_idx = matching_idx

    # --- 2. SECTIONS AND ITEMS CHECKS ---
    j_sections = jekyll_data.get("sections", [])
    if not isinstance(j_sections, list):
        j_sections = []

    last_sec_idx = -1
    for doc_sec in doc_data.get("sections", []):
        sec_title = doc_sec["title"]
        norm_title = normalize_text(sec_title)

        # Match Jekyll section
        matched_j_sec = None
        matched_j_idx = None
        for j_idx, s in enumerate(j_sections):
            if not isinstance(s, dict):
                continue
            stitle = s.get("title")
            if not stitle:
                continue
            j_norm = normalize_text(str(stitle))
            if norm_title == j_norm or (len(j_norm) > 4 and j_norm in norm_title) or (len(norm_title) > 4 and norm_title in j_norm) or text_similarity(norm_title, j_norm) >= 0.85:
                matched_j_sec = s
                matched_j_idx = j_idx
                break

        # Special case: Key Benefits is inside Our Solutions section in Jekyll
        if not matched_j_sec and "benefit" in norm_title:
            for j_idx, s in enumerate(j_sections):
                if isinstance(s, dict) and "solution" in normalize_text(str(s.get("title", ""))):
                    matched_j_sec = s
                    matched_j_idx = j_idx
                    break

        total_checks += 1
        if not matched_j_sec:
            discrepancies.append({
                "type": "MISSING_SECTION",
                "section": sec_title,
                "message": f"Section '{sec_title}' is completely missing from Jekyll YAML Front Matter."
            })
            continue

        passed_checks += 1

        # Check section sequence order
        if matched_j_idx is not None and "benefit" not in norm_title:
            total_checks += 1
            if matched_j_idx < last_sec_idx:
                discrepancies.append({
                    "type": "ORDER_MISMATCH",
                    "section": sec_title,
                    "message": f"Section '{sec_title}' is out of order in Jekyll front matter."
                })
            else:
                passed_checks += 1
            last_sec_idx = matched_j_idx

        # --- A. Frequently Asked Questions (FAQs) ---
        if doc_sec.get("faqs"):
            d_faqs = doc_sec["faqs"]
            j_faqs = matched_j_sec.get("items", [])
            last_f_idx = -1
            matched_f_indices = set()
            for i, d_faq in enumerate(d_faqs):
                total_checks += 1 # Question match check
                dq = d_faq.get("question", "")
                da = d_faq.get("answer", "")
                norm_da = normalize_text(da)

                m_idx = None
                if i < len(j_faqs) and match_titles(dq, j_faqs[i].get("question", "")):
                    m_idx = i
                else:
                    for jf_idx, jf in enumerate(j_faqs):
                        if jf_idx not in matched_f_indices and match_titles(dq, jf.get("question", "")):
                            m_idx = jf_idx
                            break

                if m_idx is None:
                    discrepancies.append({
                        "type": "MISSING_FAQ",
                        "section": sec_title,
                        "question": dq,
                        "message": f"FAQ '{dq}' missing in '{sec_title}'."
                    })
                    continue

                matched_f_indices.add(m_idx)
                passed_checks += 1
                jf_item = j_faqs[m_idx]
                ja = jf_item.get("answer", "")
                norm_ja = normalize_text(ja)

                # Order check
                total_checks += 1
                if m_idx < last_f_idx or m_idx != i:
                    discrepancies.append({
                        "type": "ORDER_MISMATCH",
                        "section": sec_title,
                        "question": dq,
                        "message": f"FAQ '{dq}' is out of order (expected position #{i+1}, found at #{m_idx+1})."
                    })
                else:
                    passed_checks += 1
                last_f_idx = m_idx

                # Answer check
                total_checks += 1
                if norm_da:
                    if not norm_ja:
                        discrepancies.append({
                            "type": "MISSING_FAQ_ANSWER",
                            "section": sec_title,
                            "question": dq,
                            "message": f"Answer missing for FAQ '{dq}' in '{sec_title}'."
                        })
                    else:
                        sim = text_similarity(norm_da, norm_ja)
                        len_ratio = min(len(norm_da), len(norm_ja)) / max(len(norm_da), len(norm_ja))
                        if norm_da == norm_ja or (sim >= 0.95 and len_ratio >= 0.90):
                            passed_checks += 1
                        else:
                            discrepancies.append({
                                "type": "FAQ_ANSWER_MISMATCH",
                                "section": sec_title,
                                "question": dq,
                                "expected": da[:80] + "...",
                                "found": ja[:80] + "..." if ja else "(empty)",
                                "message": f"Answer mismatch for FAQ '{dq}' in '{sec_title}'."
                            })
                else:
                    passed_checks += 1

            for jf_idx, jf in enumerate(j_faqs):
                if jf_idx not in matched_f_indices:
                    total_checks += 1
                    discrepancies.append({
                        "type": "EXTRA_FAQ",
                        "section": sec_title,
                        "question": jf.get("question", f"FAQ #{jf_idx+1}"),
                        "message": f"Extra FAQ '{jf.get('question')}' found in Jekyll '{sec_title}'."
                    })

        # --- B. Technology Rows ---
        if doc_sec.get("table_rows"):
            d_rows = doc_sec["table_rows"]
            j_rows = matched_j_sec.get("rows", [])
            last_r_idx = -1
            matched_r_indices = set()
            for i, d_row in enumerate(d_rows):
                total_checks += 1 # Label match check
                dlbl = d_row.get("label", "")
                dval = d_row.get("value", "")
                norm_dval = normalize_text(dval)

                m_idx = None
                if i < len(j_rows) and match_titles(dlbl, j_rows[i].get("label", "")):
                    m_idx = i
                else:
                    for jr_idx, jr in enumerate(j_rows):
                        if jr_idx not in matched_r_indices and match_titles(dlbl, jr.get("label", "")):
                            m_idx = jr_idx
                            break

                if m_idx is None:
                    discrepancies.append({
                        "type": "MISSING_TECH",
                        "section": sec_title,
                        "label": dlbl,
                        "message": f"Technology row '{dlbl}' missing in '{sec_title}'."
                    })
                    continue

                matched_r_indices.add(m_idx)
                passed_checks += 1
                jr_item = j_rows[m_idx]
                jval = jr_item.get("value", "")
                norm_jval = normalize_text(jval)

                # Order check
                total_checks += 1
                if m_idx < last_r_idx or m_idx != i:
                    discrepancies.append({
                        "type": "ORDER_MISMATCH",
                        "section": sec_title,
                        "label": dlbl,
                        "message": f"Technology '{dlbl}' is out of order (expected position #{i+1}, found at #{m_idx+1})."
                    })
                else:
                    passed_checks += 1
                last_r_idx = m_idx

                # Value check
                total_checks += 1
                if norm_dval:
                    sim = text_similarity(norm_dval, norm_jval)
                    if norm_dval == norm_jval or sim >= 0.90:
                        passed_checks += 1
                    else:
                        discrepancies.append({
                            "type": "TECH_VALUE_MISMATCH",
                            "section": sec_title,
                            "label": dlbl,
                            "expected": dval,
                            "found": jval,
                            "message": f"Value mismatch for technology '{dlbl}'. Expected: '{dval}' | Found: '{jval}'."
                        })
                else:
                    passed_checks += 1

            for jr_idx, jr in enumerate(j_rows):
                if jr_idx not in matched_r_indices:
                    total_checks += 1
                    discrepancies.append({
                        "type": "EXTRA_TECH",
                        "section": sec_title,
                        "label": jr.get("label", f"Row #{jr_idx+1}"),
                        "message": f"Extra technology row '{jr.get('label')}' found in Jekyll '{sec_title}'."
                    })

        # --- C. Section Items ---
        if doc_sec.get("items"):
            d_items = doc_sec["items"]

            # 1. Business Results Metrics
            if "business results" in norm_title or (d_items and isinstance(d_items[0], dict) and "value" in d_items[0]):
                j_items = matched_j_sec.get("items", [])
                last_res_idx = -1
                matched_res_indices = set()
                for i, d_res in enumerate(d_items):
                    total_checks += 1 # Metric value match
                    d_val = normalize_text(str(d_res.get("value", "")))
                    d_txt = str(d_res.get("text", "")).strip()
                    d_sub = str(d_res.get("subtext", "")).strip()

                    m_idx = None
                    if i < len(j_items) and d_val == normalize_text(str(j_items[i].get("value", ""))):
                        m_idx = i
                    else:
                        for jr_idx, jr in enumerate(j_items):
                            if jr_idx not in matched_res_indices and d_val == normalize_text(str(jr.get("value", ""))):
                                m_idx = jr_idx
                                break

                    if m_idx is None:
                        discrepancies.append({
                            "type": "MISSING_RESULT",
                            "section": sec_title,
                            "metric": f"{d_res.get('value')} {d_txt}",
                            "message": f"Result metric '{d_res.get('value')} {d_txt}' missing in '{sec_title}'."
                        })
                        continue

                    matched_res_indices.add(m_idx)
                    passed_checks += 1
                    jr_item = j_items[m_idx]
                    j_txt = str(jr_item.get("text", "")).strip()
                    j_sub = str(jr_item.get("subtext", "")).strip()

                    # Order check
                    total_checks += 1
                    if m_idx < last_res_idx or m_idx != i:
                        discrepancies.append({
                            "type": "ORDER_MISMATCH",
                            "section": sec_title,
                            "metric": d_res.get("value"),
                            "message": f"Result metric '{d_res.get('value')}' is out of order (expected position #{i+1}, found at #{m_idx+1})."
                        })
                    else:
                        passed_checks += 1
                    last_res_idx = m_idx

                    # Text check
                    total_checks += 1
                    norm_dtxt = normalize_text(d_txt)
                    norm_jtxt = normalize_text(j_txt)
                    if norm_dtxt == norm_jtxt or text_similarity(norm_dtxt, norm_jtxt) >= 0.85:
                        passed_checks += 1
                    else:
                        discrepancies.append({
                            "type": "RESULT_TEXT_MISMATCH",
                            "section": sec_title,
                            "metric": d_res.get("value"),
                            "expected": d_txt,
                            "found": j_txt,
                            "message": f"Text mismatch for result metric '{d_res.get('value')}'. Expected: '{d_txt}' | Found: '{j_txt}'."
                        })

                    # Subtext check
                    if d_sub:
                        total_checks += 1
                        norm_dsub = normalize_text(d_sub)
                        norm_jsub = normalize_text(j_sub)
                        if norm_dsub == norm_jsub or text_similarity(norm_dsub, norm_jsub) >= 0.80 or norm_dsub in norm_jsub or norm_jsub in norm_dsub:
                            passed_checks += 1
                        else:
                            discrepancies.append({
                                "type": "RESULT_SUBTEXT_MISMATCH",
                                "section": sec_title,
                                "metric": d_res.get("value"),
                                "expected": d_sub[:80] + "...",
                                "found": j_sub[:80] + "..." if j_sub else "(empty)",
                                "message": f"Subtext mismatch for result metric '{d_res.get('value')}'. Expected: '{d_sub[:80]}...' | Found: '{j_sub[:80]}...'."
                            })

                for jr_idx, jr in enumerate(j_items):
                    if jr_idx not in matched_res_indices:
                        total_checks += 1
                        discrepancies.append({
                            "type": "EXTRA_RESULT",
                            "section": sec_title,
                            "metric": f"{jr.get('value')} {jr.get('text', '')}",
                            "message": f"Extra result metric '{jr.get('value')}' found in Jekyll '{sec_title}'."
                        })

            # 2. Challenges or Key Features (dict with title & description)
            elif d_items and isinstance(d_items[0], dict) and "title" in d_items[0]:
                j_items = matched_j_sec.get("items", [])
                last_it_idx = -1
                matched_it_indices = set()
                for i, d_item in enumerate(d_items):
                    total_checks += 1 # Title match check
                    d_title = d_item.get("title", "")
                    d_desc = d_item.get("description", "")
                    norm_d_desc = normalize_text(d_desc)

                    m_idx = None
                    if i < len(j_items) and match_titles(d_title, j_items[i].get("title", "")):
                        m_idx = i
                    else:
                        for jit_idx, jit in enumerate(j_items):
                            if jit_idx not in matched_it_indices and match_titles(d_title, jit.get("title", "")):
                                m_idx = jit_idx
                                break

                    if m_idx is None:
                        discrepancies.append({
                            "type": "MISSING_ITEM",
                            "section": sec_title,
                            "item": d_title,
                            "message": f"Item '{d_title}' missing in '{sec_title}'."
                        })
                        continue

                    matched_it_indices.add(m_idx)
                    passed_checks += 1
                    jit_item = j_items[m_idx]
                    j_desc = jit_item.get("description", "")
                    norm_j_desc = normalize_text(j_desc)

                    # Order check
                    total_checks += 1
                    if m_idx < last_it_idx or m_idx != i:
                        discrepancies.append({
                            "type": "ORDER_MISMATCH",
                            "section": sec_title,
                            "item": d_title,
                            "message": f"Item '{d_title}' in '{sec_title}' is out of order (expected position #{i+1}, found at #{m_idx+1})."
                        })
                    else:
                        passed_checks += 1
                    last_it_idx = m_idx

                    # Description / Sublabel check
                    total_checks += 1
                    if norm_d_desc:
                        if not norm_j_desc:
                            discrepancies.append({
                                "type": "MISSING_DESCRIPTION",
                                "section": sec_title,
                                "item": d_title,
                                "message": f"Description missing for item '{d_title}' in '{sec_title}'."
                            })
                        else:
                            sim = text_similarity(norm_d_desc, norm_j_desc)
                            len_ratio = min(len(norm_d_desc), len(norm_j_desc)) / max(len(norm_d_desc), len(norm_j_desc))
                            if norm_d_desc == norm_j_desc or (sim >= 0.95 and len_ratio >= 0.90):
                                passed_checks += 1
                            else:
                                discrepancies.append({
                                    "type": "DESCRIPTION_MISMATCH",
                                    "section": sec_title,
                                    "item": d_title,
                                    "expected": d_desc[:80] + "...",
                                    "found": j_desc[:80] + "..." if j_desc else "(empty)",
                                    "message": f"Description mismatch for item '{d_title}' in '{sec_title}'."
                                })
                    else:
                        passed_checks += 1

                for jit_idx, jit in enumerate(j_items):
                    if jit_idx not in matched_it_indices:
                        total_checks += 1
                        discrepancies.append({
                            "type": "EXTRA_ITEM",
                            "section": sec_title,
                            "item": jit.get("title", f"Item #{jit_idx+1}"),
                            "message": f"Extra item '{jit.get('title')}' found in Jekyll '{sec_title}'."
                        })

            # 3. Simple Bullet Points (str)
            else:
                j_raw_content = str(matched_j_sec.get("content", "") or matched_j_sec.get("description", ""))
                if matched_j_sec.get("points"):
                    j_raw_content += "\n" + "\n".join(str(p) for p in matched_j_sec.get("points", []))
                if matched_j_sec.get("items"):
                    j_raw_content += "\n" + "\n".join(
                        f"{it.get('title', '')} {it.get('description', '')}" if isinstance(it, dict)
                        else str(it)
                        for it in matched_j_sec.get("items", [])
                    )
                norm_j_raw = normalize_text(j_raw_content)

                last_pos = -1
                for i, d_bullet in enumerate(d_items):
                    total_checks += 1
                    d_b_norm = normalize_text(d_bullet)
                    pos = norm_j_raw.find(d_b_norm)
                    if pos != -1:
                        total_checks += 1
                        if pos < last_pos:
                            discrepancies.append({
                                "type": "ORDER_MISMATCH",
                                "section": sec_title,
                                "item": d_bullet[:40] + "...",
                                "message": f"Bullet point '{d_bullet[:40]}...' in '{sec_title}' is out of order."
                            })
                        else:
                            passed_checks += 1
                        passed_checks += 1
                        last_pos = pos
                    else:
                        sim_found = False
                        for line in j_raw_content.splitlines():
                            if text_similarity(d_b_norm, normalize_text(line)) >= 0.85:
                                sim_found = True
                                break
                        if sim_found:
                            passed_checks += 1
                        else:
                            discrepancies.append({
                                "type": "MISSING_BULLET",
                                "section": sec_title,
                                "item": d_bullet[:60] + "...",
                                "message": f"Bullet point '{d_bullet[:60]}...' missing in '{sec_title}'."
                            })

        # --- D. Paragraphs ---
        if doc_sec.get("paragraphs"):
            j_raw_content = str(matched_j_sec.get("content", "") or matched_j_sec.get("description", ""))
            if not j_raw_content and matched_j_sec.get("items"):
                j_raw_content = "\n\n".join(
                    f"{it.get('title', '')} {it.get('description', '')}" if isinstance(it, dict)
                    else str(it)
                    for it in matched_j_sec.get("items", [])
                )
            norm_j_raw = normalize_text(j_raw_content)
            last_p_pos = -1
            for p in doc_sec["paragraphs"]:
                if len(p.strip()) > 30:
                    total_checks += 1
                    norm_p = normalize_text(p)
                    p_pos = norm_j_raw.find(norm_p)
                    if p_pos != -1:
                        total_checks += 1
                        if p_pos < last_p_pos:
                            discrepancies.append({
                                "type": "ORDER_MISMATCH",
                                "section": sec_title,
                                "text": p[:60] + "...",
                                "message": f"Paragraph in '{sec_title}' is out of order: '{p[:60]}...'"
                            })
                        else:
                            passed_checks += 1
                        passed_checks += 1
                        last_p_pos = p_pos
                    elif text_similarity(norm_p, norm_j_raw) >= 0.80 or any(text_similarity(norm_p, normalize_text(line)) >= 0.80 for line in j_raw_content.split("\n\n")):
                        passed_checks += 1
                    else:
                        discrepancies.append({
                            "type": "MISSING_PARAGRAPH",
                            "section": sec_title,
                            "text": p[:70] + "...",
                            "message": f"Paragraph in '{sec_title}' missing or modified: '{p[:70]}...'"
                        })

    match_score = (passed_checks / total_checks * 100.0) if total_checks > 0 else 100.0
    return match_score, discrepancies, success_items

# --- AGENT AUTOMATION WITH SESSION PERSISTENCE ---

def assign_agent_content_task(
    local_file: str,
    doc_path: str,
    discrepancies: List[Dict[str, str]],
    jekyll_dir: str = JEKYLL_PROJECT_DIR,
    mode: str = "auto",
    conversation_id: Optional[str] = None
) -> Tuple[bool, Optional[str]]:
    """
    Assigns task to Antigravity CLI (agy) to automatically synchronize the Jekyll Front Matter
    with the Google Doc markdown specification.
    Reuses existing conversation session for speed and shared context.
    """
    disc_summary = "\n".join([f"- [{d['type']}] {d.get('message', '')}" for d in discrepancies[:15]])
    if len(discrepancies) > 15:
        disc_summary += f"\n... and {len(discrepancies) - 15} more items."

    prompt = (
        f"@{local_file}\n"
        f"In the project Jekyll markdown file, synchronize all content in the YAML Front Matter as per the Google Doc specification - @{doc_path}\n\n"
        f"Discrepancies detected:\n{disc_summary}\n\n"
        f"Instructions:\n"
        f"1. In the target file ({local_file}), locate the YAML Front Matter (between the first and second --- delimiters).\n"
        f"2. Add or update all missing/mismatched content (sections, items, FAQs, technology rows, business results, metadata) so that it strictly matches @{doc_path}.\n"
        f"3. Maintain proper Jekyll YAML structure (e.g. sections list with type, title, items, content, rows).\n"
        f"4. Preserve existing root-relative internal links or convert URLs to root-relative paths like /service-name.\n"
        f"5. Apply the edits directly to {local_file}.\n"
        f"6. IMPORTANT OUTPUT RULE: When finished, respond ONLY with the exact single line below and NOTHING else:\n"
        f"All content discrepancies have been resolved in {local_file}."
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

    print(f"🤖 Assigning content sync task to agent ({session_info})...")

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
                print(f"✅ Agent applied edits to {local_file}.\n")
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
                print(f"✅ Agent applied edits to {local_file}.\n")

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
    skip_rows: Optional[List[int]] = None,
    threshold: float = DEFAULT_SIMILARITY_THRESHOLD
):
    state = load_progress()
    active_skip_rows = set(skip_rows if skip_rows is not None else ROWS_TO_SKIP)

    # Determine starting row
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

    active_conv_id = conversation_id or state.get("conversation_id")
    if active_conv_id:
        print(f"🔗 Reusing agent session: {active_conv_id}")

    sheets, drive, docs = get_services()
    resolved_sheet = get_resolved_sheet_name(sheets, SPREADSHEET_ID, sheet_name)

    # Fetch rows
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

                old_url = row[1] if len(row) > 1 else ""
                doc_col = row[2] if len(row) > 2 else ""
                new_url = row[4] if len(row) > 4 else ""

                # Extract Doc ID
                doc_id = extract_doc_id(doc_col)
                if not doc_id:
                    for cell in row:
                        doc_id = extract_doc_id(cell)
                        if doc_id:
                            break

                if not doc_id:
                    empty_docs_skipped += 1
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    continue

                if empty_docs_skipped > 0:
                    print(f"⏩ Skipped {empty_docs_skipped} row(s) without Google Doc links.")
                    empty_docs_skipped = 0

                actionable_count += 1
                live_link = new_url if (new_url and new_url.startswith("http")) else old_url

                if live_link.startswith("http"):
                    slug_part = urlparse(live_link).path.strip("/").split("/")[-1]
                    page_name = slug_part.replace("-", " ").title() if slug_part else f"Row {current_row_num}"
                else:
                    page_name = old_url if old_url else f"Row {current_row_num}"

                batch_info = f"[{actionable_count}/{batch_target}] " if batch_target else ""
                print(f"--- Verifying Content {batch_info}[Row {current_row_num}] {page_name} ---")

                # 1. Download & Parse Google Doc (using Docs API with Drive export fallback)
                doc_data = None
                doc_text = ""
                try:
                    doc_obj = docs.documents().get(documentId=doc_id).execute()
                    doc_data, doc_text = parse_docs_api(doc_obj)
                except Exception:
                    try:
                        doc_text = drive.files().export(fileId=doc_id, mimeType="text/plain").execute().decode("utf-8")
                        doc_data = parse_doc_content(doc_text)
                    except Exception as e:
                        print(f"⚠️ Could not download Google Doc ({doc_id}): {e}\n")
                        update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                        row_idx += 1
                        if batch_target and actionable_count >= batch_target:
                            break
                        continue

                # 2. Find local Jekyll file
                local_file = find_matching_jekyll_file(page_name, live_link, jekyll_dir=jekyll_dir, folders=folders)
                if not local_file:
                    print(f"⚠️ Could not find local .md file for '{page_name}' (Slug: {urlparse(live_link).path}).\n")
                    update_progress(last_processed_row=current_row_num, next_row=current_row_num + 1, sheet_name=resolved_sheet, conversation_id=active_conv_id)
                    row_idx += 1
                    if batch_target and actionable_count >= batch_target:
                        break
                    continue

                # 3. Parse Jekyll Front Matter
                jekyll_data = parse_jekyll_front_matter(local_file)

                # 4. Compare Content
                score, discrepancies, _ = compare_content(doc_data, jekyll_data, threshold=threshold)

                if not discrepancies:
                    print(f"✅ Content Verified! (Match Score: {score:.1f}%) in `{local_file}`.\n")
                    update_progress(
                        last_processed_row=current_row_num,
                        next_row=current_row_num + 1,
                        sheet_name=resolved_sheet,
                        conversation_id=active_conv_id
                    )
                    row_idx += 1
                    if batch_target and actionable_count >= batch_target:
                        break
                    continue

                print(f"❌ Found {len(discrepancies)} discrepancy(ies) (Match Score: {score:.1f}%):")
                for d in discrepancies[:8]:
                    print(f"   - {d.get('message', '')}")
                if len(discrepancies) > 8:
                    print(f"   ... and {len(discrepancies) - 8} more.")

                # 5. Prepare doc markdown in _work_docs
                work_docs_dir = os.path.join(jekyll_dir, "_work_docs")
                clean_name = re.sub(r'[\/:*?"<>|]', "_", page_name).strip()
                try:
                    os.makedirs(work_docs_dir, exist_ok=True)
                    doc_path = os.path.join(work_docs_dir, f"{clean_name}.md")
                except Exception:
                    doc_path = os.path.join(tempfile.gettempdir(), f"{clean_name}.md")

                with open(doc_path, "w", encoding="utf-8") as f:
                    f.write(doc_text)

                # 6. Dispatch to Agent if mode != 'skip'
                if agent_mode == "skip":
                    print("⏩ Skipping agent task as requested.\n")
                else:
                    _, active_conv_id = assign_agent_content_task(
                        local_file=local_file,
                        doc_path=doc_path,
                        discrepancies=discrepancies,
                        jekyll_dir=jekyll_dir,
                        mode=agent_mode,
                        conversation_id=active_conv_id
                    )
                    # Re-verify after agent execution
                    updated_jekyll = parse_jekyll_front_matter(local_file)
                    new_score, remaining_disc, _ = compare_content(doc_data, updated_jekyll, threshold=threshold)
                    if not remaining_disc:
                        print(f"🎉 All content discrepancies resolved! (Final Score: 100%)\n")
                    else:
                        print(f"ℹ️ Post-agent score: {new_score:.1f}% ({len(remaining_disc)} remaining item(s)).\n")

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

            if row_idx >= total_rows:
                print(f"\n🎉 All {total_rows} rows in '{resolved_sheet}' have been processed!")
                break

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
    parser = argparse.ArgumentParser(description="Batch Front Matter Content Verifier")
    parser.add_argument("--sheet", type=str, default=DEFAULT_SHEET_NAME, help=f"Sheet/tab name to process (default: '{DEFAULT_SHEET_NAME}')")
    parser.add_argument("--start", type=int, default=None, help=f"Starting row number (default: resume from {PROGRESS_FILE} or {DEFAULT_START_ROW})")
    parser.add_argument("--batch", type=int, default=DEFAULT_BATCH_SIZE, help="Batch size (number of actionable rows to process; 0 for all remaining)")
    parser.add_argument("--jekyll-dir", type=str, default=JEKYLL_PROJECT_DIR, help="Absolute path to Jekyll project directory")
    parser.add_argument("--folders", nargs="+", default=CONTENT_FOLDERS, help="Content folders to scan (e.g. _services _technologies _work)")
    parser.add_argument(
        "--agent-mode",
        choices=["auto", "interactive", "skip"],
        default="auto",
        help="Agent execution mode: 'auto' (direct automated edits, default), 'interactive', or 'skip'"
    )
    parser.add_argument("--conversation", type=str, default=None, help="Explicit Antigravity conversation ID to resume")
    parser.add_argument("--new-session", action="store_true", help="Start a new agent conversation session instead of resuming")
    parser.add_argument("--skip-rows", nargs="+", type=int, default=ROWS_TO_SKIP, help=f"Row numbers to skip (default: {ROWS_TO_SKIP})")
    parser.add_argument("--threshold", type=float, default=DEFAULT_SIMILARITY_THRESHOLD, help=f"Similarity threshold (default: {DEFAULT_SIMILARITY_THRESHOLD})")
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
        skip_rows=args.skip_rows,
        threshold=args.threshold
    )
