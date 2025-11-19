import json
import os
import re
from decimal import Decimal


LOG_DIR = "log"


DEFAULTS = {
    # --- Fees tab defaults ---
    "fees_search_token": "KFB",  # typed in file-search box (once at start)
    "fees_bad_prefixes": "SVRAGS;SVR-AGS;Skrags;SV RAGS",  # semicolon-separated
    "fees_overlay_skip_waits": True,  # behave like Streitwert overlay-only waits
    "fees_file_search_region": [0.10, 0.08, 0.25, 0.05],  # where we click+type KFB
    "fees_seiten_region": [0.08, 0.84, 0.84, 0.12],  # thumbnails strip
    "fees_pages_max_clicks": 12,  # safety upper bound of page clicks
    "fees_csv_path": "fees_results.csv",  # output
    "rdp_title_regex": r".* - Remote Desktop Connection",
    "excel_path": "input.xlsx",
    "excel_sheet": "Sheet1",
    "input_column": "query",
    "results_csv": "rdp_results.csv",
    "tesseract_path": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    "tesseract_lang": "deu+eng",
    "type_delay": 0.02,
    "post_search_wait": 1.2,
    "search_point": [0.25, 0.10],  # relative x,y within RDP client
    "result_region": [0.15, 0.20, 0.80, 0.35],  # relative l,t,w,h within RDP client
    "start_cell": "",  # e.g., "B2"
    "max_rows": 0,
    "line_band_px": 40,  # used only for manual-line OCR fallback
    "picked_line_rel_y": None,
    "typing_test_text": "TEST123",
    "line_offset_px": 0,
    "upscale_x": 4,
    "color_ocr": True,
    "auto_green": True,
    # NEW: full-region parsing (works even when no green row is selected)
    "use_full_region_parse": True,
    "keyword": "Honorar",
    "normalize_ocr": True,
    # -------- NEW: Amount Region Profiles --------
    # Each profile is stored relative to "result_region"
    # { "name": str, "keyword": str, "sub_region": [l, t, w, h] }
    "amount_profiles": [],  # list of dicts
    "active_amount_profile": "",  # profile name
    "use_amount_profile": False,  # if True, restrict OCR to profile sub-region
    # --- Streitwert workflow (NEW) ---
    "doclist_region": [0.10, 0.24, 0.78, 0.50],  # list/table area with documents
    "pdf_search_point": [0.55, 0.10],  # the PDF viewer's search field
    "pdf_hits_point": [0.08, 0.32],  # button inside the PDF hits pane
    "pdf_hits_second_point": [0.08, 0.40],  # optional 2nd PDF result button
    "pdf_hits_third_point": [0.08, 0.48],  # optional 3rd PDF result button
    "pdf_text_region": [0.20, 0.18, 0.74, 0.68],  # main page text area
    "includes": "Urt,SWB,SW",  # rows to include if they contain any of these
    "excludes": "SaM,KLE",  # rows to skip if they contain any of these
    "exclude_prefix_k": True,  # also skip rows starting with 'K' (e.g. 'K9 Urteil')
    "streitwert_term": "Streitwert",  # term to type into PDF search
    "streitwert_results_csv": "streitwert_results.csv",
    "doc_open_wait": 1.2,  # wait (s) after opening a doc
    "pdf_hit_wait": 1.0,  # wait (s) after clicking a search hit
    "pdf_view_extra_wait": 2.0,  # wait (s) after pressing the PDF results button
    "doc_view_point": [0.88, 0.12],  # "View" button to open the selected doc
    "pdf_close_point": [0.97, 0.05],  # close button for the PDF viewer window
    "streitwert_overlay_skip_waits": False,  # rely solely on overlay detection delays
    "ignore_top_doc_row": False,  # skip the first/topmost Streitwert match
    # --- Rechnungen workflow (NEW) ---
    "rechnungen_region": [0.55, 0.30, 0.35, 0.40],
    "rechnungen_results_csv": "Streitwert_Results_Rechnungen.csv",
    "rechnungen_only_results_csv": "rechnungen_only_results.csv",
    "rechnungen_gg_region": [0.55, 0.30, 0.35, 0.40],
    "rechnungen_gg_results_csv": "rechnungen_gg_results.csv",
    "rechnungen_search_wait": 1.2,
    "rechnungen_region_wait": 0.8,
    "rechnungen_overlay_skip_waits": False,
    "log_folder": LOG_DIR,
    "log_extract_results_csv": "streitwert_log_extract.csv",
    # New: AZ Instanz table detection
    "instance_region": [0.10, 0.10, 0.30, 0.10],  # (l,t,w,h) relative to RDP window
    "instance_row_rel_top": 0.45,  # start scanning at 45% from top (below headers)
}

CFG_FILE = "rdp_automation_config.json"

STREITWERT_MIN_AMOUNT = Decimal("1000")

FORCED_STREITWERT_EXCLUDES = [
    ("KAA GS", re.compile(r"\bKAA(?:\s+|-)?GS\b", re.IGNORECASE)),
    ("KAAGS", re.compile(r"\bKAAGS\b", re.IGNORECASE)),
    ("KAA", re.compile(r"\bKAA\b", re.IGNORECASE)),
    ("KFB(vA)", re.compile(r"\bKFB\s*\(vA\)\b", re.IGNORECASE)),
    ("KFB vA", re.compile(r"\bKFB\s*vA\b", re.IGNORECASE)),
    ("KFB", re.compile(r"\bKFB\b", re.IGNORECASE)),
    ("KLE", re.compile(r"\bKLE\b", re.IGNORECASE)),
    ("GS", re.compile(r"\bGS\b", re.IGNORECASE)),
]


def ensure_log_dir():
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
    except Exception:
        pass


def sanitize_filename(value: str) -> str:
    if not value:
        return "ocr_log"
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value))
    safe = safe.strip("._-")
    if len(safe) > 120:
        safe = safe[:120]
    return safe or "ocr_log"


def load_cfg():
    if os.path.exists(CFG_FILE):
        with open(CFG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        cfg = DEFAULTS.copy()
        cfg.update(data)
        cfg.setdefault("amount_profiles", [])
        cfg.setdefault("active_amount_profile", "")
        cfg.setdefault("use_amount_profile", False)
        return cfg
    return DEFAULTS.copy()


def save_cfg(cfg):
    with open(CFG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
