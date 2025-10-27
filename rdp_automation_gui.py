import os, io, json, time, threading, re, unicodedata
from decimal import Decimal, InvalidOperation
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from PIL import (
    Image,
    ImageTk,
    ImageFilter,
    ImageOps,
    ImageStat,
    ImageDraw,
    ImageEnhance,
)
import pandas as pd
import pyautogui
from pywinauto import Desktop
from mss import mss
import pytesseract
import numpy as np

# Try OpenCV (optional; used for green-row detection; app works without it)
try:
    import cv2

    _HAS_CV2 = True
except Exception:
    _HAS_CV2 = False

# ------------------ Defaults & Config ------------------
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


# ------------------ Helpers ------------------
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


def connect_rdp_window(title_re):
    win = Desktop(backend="uia").window(title_re=title_re)
    win.wait("exists ready", timeout=10)
    try:
        client = win.child_window(control_type="Pane")
        r = client.rectangle()
    except Exception:
        r = win.rectangle()
        r = type(
            "Rect",
            (),
            {
                "left": r.left + 8,
                "top": r.top + 40,
                "right": r.right - 8,
                "bottom": r.bottom - 8,
            },
        )()
    win.set_focus()
    return win, (r.left, r.top, r.right, r.bottom)


def rel_to_abs(rect, rel_box):
    left, top, right, bottom = rect
    w, h = right - left, bottom - top
    if len(rel_box) == 2:
        rx, ry = rel_box
        return int(left + rx * w), int(top + ry * h)
    rl, rt, rw, rh = rel_box
    return (int(left + rl * w), int(top + rt * h), int(rw * w), int(rh * h))


def abs_to_rel(rect, abs_point=None, abs_box=None):
    left, top, right, bottom = rect
    w, h = right - left, bottom - top
    if abs_point:
        x, y = abs_point
        return [(x - left) / w, (y - top) / h]
    x, y, bw, bh = abs_box
    return [(x - left) / w, (y - top) / h, bw / w, bh / h]


# Global thread-local storage for MSS
_thread_local = threading.local()


def get_mss():
    if not hasattr(_thread_local, "sct"):
        _thread_local.sct = mss()
    return _thread_local.sct


def grab_xywh(x, y, w, h):
    sct = get_mss()
    shot = sct.grab({"left": x, "top": y, "width": w, "height": h})
    return Image.frombytes("RGB", shot.size, shot.rgb)


def upscale_pil(img_pil, scale=3):
    return (
        img_pil.resize((img_pil.width * scale, img_pil.height * scale), Image.LANCZOS)
        if scale > 1
        else img_pil
    )


def do_ocr_color(img, lang="eng", psm=6):
    common = r"--oem 3 -c preserve_interword_spaces=1"
    cfg = f"--psm {psm} {common}"
    return pytesseract.image_to_string(img, lang=lang, config=cfg).strip()


def do_ocr_data(img, lang="eng", psm=6):
    """Return pytesseract TSV DataFrame for line-by-line parsing."""
    common = r"--oem 3 -c preserve_interword_spaces=1"
    cfg = f"--psm {psm} {common}"
    return pytesseract.image_to_data(
        img, lang=lang, config=cfg, output_type=pytesseract.Output.DATAFRAME
    )


# ---------- Green-row detection (optional) ----------
def find_green_band(color_img_pil):
    if not _HAS_CV2:
        return None
    img = np.array(color_img_pil)  # RGB
    hsv = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
    lower = np.array([35, 40, 40], dtype=np.uint8)
    upper = np.array([85, 255, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    best, best_score = None, -1
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        score = w * max(1, h) - 3 * h
        if score > best_score:
            best_score, best = score, (x, y, w, h)
    if best is None:
        return None
    x, y, w, h = best
    pad_x = max(6, w // 20)
    pad_y = max(2, h // 6)
    x0 = max(0, x - pad_x)
    y0 = max(0, y - pad_y)
    x1 = min(img.shape[1], x + w + pad_x)
    y1 = min(img.shape[0], y + h + pad_y)
    return color_img_pil.crop((x0, y0, x1, y1))


# ---------- Doclist OCR + Click helpers (used by Fees/Streitwert) ----------


def _ocr_doclist_rows(self):
    """
    OCR the calibrated doclist_region and return a list of row TEXTS (strings).
    Uses Tesseract's line grouping when available; otherwise groups by Y.
    """
    if not self._has("doclist_region"):
        self.log_print("[Doclist OCR] doclist_region not configured.")
        return []

    x, y, w, h = self._get("doclist_region")
    img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
    df = do_ocr_data(img, lang=(self.lang_var.get().strip() or "deu+eng"), psm=6)

    if df is None or "text" not in df.columns:
        return []

    # Keep only meaningful tokens
    def _ok(s):
        return bool(s) and str(s).strip() not in ("", "nan", None)

    df = df.copy()
    df["text"] = df["text"].astype(str)
    df = df[df["text"].apply(_ok)]
    if df.empty:
        return []

    rows = []
    if {"block_num", "par_num", "line_num", "left", "top", "width", "height"}.issubset(
        df.columns
    ):
        # Group by Tesseract line identifiers
        for (b, p, l), g in df.groupby(["block_num", "par_num", "line_num"], sort=True):
            g = g.sort_values("left")
            txt = " ".join(t.strip() for t in g["text"].tolist() if t.strip())
            if txt:
                rows.append(txt)
    else:
        # Fallback: group by Y proximity
        df = df.sort_values("top")
        tol = max(8, int(img.height * 0.01))
        current_y = None
        buf = []
        for _, r in df.iterrows():
            if current_y is None or abs(int(r["top"]) - current_y) <= tol:
                buf.append(str(r["text"]).strip())
                if current_y is None:
                    current_y = int(r["top"])
            else:
                line = " ".join(t for t in buf if t)
                if line:
                    rows.append(line)
                buf = [str(r["text"]).strip()]
                current_y = int(r["top"])
        if buf:
            line = " ".join(t for t in buf if t)
            if line:
                rows.append(line)

    # Light cleanup & de-dup short artifacts
    cleaned = []
    for s in rows:
        s2 = " ".join(s.split())
        if len(s2) >= 2:
            if not cleaned or cleaned[-1] != s2:
                cleaned.append(s2)
    self.log_print(f"[Doclist OCR] lines: {len(cleaned)}")
    return cleaned


def _ocr_doclist_rows_boxes(self):
    """
    Return list of (text, (lx,ty,rx,by)) for each detected line in doclist_region.
    These boxes are **relative to the doclist image** (not absolute screen).
    """
    if not self._has("doclist_region"):
        self.log_print("[Doclist OCR] doclist_region not configured.")
        return []

    x, y, w, h = self._get("doclist_region")
    img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
    df = do_ocr_data(img, lang=(self.lang_var.get().strip() or "deu+eng"), psm=6)
    if df is None or "text" not in df.columns:
        return []

    def _ok(s):
        return bool(s) and str(s).strip() not in ("", "nan", None)

    df = df.copy()
    df["text"] = df["text"].astype(str)
    df = df[df["text"].apply(_ok)]
    if df.empty:
        return []

    lines = []
    if {"block_num", "par_num", "line_num", "left", "top", "width", "height"}.issubset(
        df.columns
    ):
        for (b, p, l), g in df.groupby(["block_num", "par_num", "line_num"], sort=True):
            g = g.sort_values("left")
            txt = " ".join(t.strip() for t in g["text"].tolist() if t.strip())
            if not txt:
                continue
            lx = int(g["left"].min())
            ty = int(g["top"].min())
            rx = int((g["left"] + g["width"]).max())
            by = int((g["top"] + g["height"]).max())
            lines.append((txt, (lx, ty, rx, by)))
    else:
        df = df.sort_values("top")
        tol = max(8, int(img.height * 0.01))
        cur_top = None
        cur = []
        for _, r in df.iterrows():
            if cur_top is None or abs(int(r["top"]) - cur_top) <= tol:
                cur.append(r)
                if cur_top is None:
                    cur_top = int(r["top"])
            else:
                if cur:
                    lx = min(int(rr["left"]) for rr in cur)
                    ty = min(int(rr["top"]) for rr in cur)
                    rx = max(int(rr["left"] + rr["width"]) for rr in cur)
                    by = max(int(rr["top"] + rr["height"]) for rr in cur)
                    txt = " ".join(
                        str(rr["text"]).strip()
                        for rr in sorted(cur, key=lambda t: int(t["left"]))
                        if str(rr["text"]).strip()
                    )
                    if txt:
                        lines.append((txt, (lx, ty, rx, by)))
                cur = [r]
                cur_top = int(r["top"])
        if cur:
            lx = min(int(rr["left"]) for rr in cur)
            ty = min(int(rr["top"]) for rr in cur)
            rx = max(int(rr["left"] + rr["width"]) for rr in cur)
            by = max(int(rr["top"] + rr["height"]) for rr in cur)
            txt = " ".join(
                str(rr["text"]).strip()
                for rr in sorted(cur, key=lambda t: int(t["left"]))
                if str(rr["text"]).strip()
            )
            if txt:
                lines.append((txt, (lx, ty, rx, by)))

    self.log_print(f"[Doclist OCR] lines+boxes: {len(lines)}")
    return lines


def _click_doclist_row(self, row_idx: int):
    """
    Click the center of the given row index inside the doclist_region using OCR boxes.
    Returns True on success, False otherwise.
    """
    if row_idx is None or row_idx < 0:
        return False
    rows = self._ocr_doclist_rows_boxes()
    if not rows or row_idx >= len(rows):
        return False

    # target box in doclist image coords
    _, (lx, ty, rx, by) = rows[row_idx]
    cx = (lx + rx) // 2
    cy = (ty + by) // 2

    # map to absolute screen coords using calibrated doclist_region and current_rect
    X, Y, W, H = self._get("doclist_region")
    abs_x = X + cx
    abs_y = Y + cy

    try:
        pyautogui.click(abs_x, abs_y)
        time.sleep(0.08)
        return True
    except Exception as e:
        self.log_print(f"[Doclist OCR] click failed: {e}")
        return False


# ---------- OCR TSV helpers (Streitwert) ----------
AMOUNT_TOKEN_TRANSLATE = str.maketrans(
    {"O": "0", "o": "0", "S": "5", "s": "5", "l": "1", "I": "1", "B": "8"}
)


# Treat punctuation, borders, and filler as noise (not real text)
_NOISE_RE = re.compile(r"^[_\-–—\|:;.,'\"`~^°()+\[\]{}<>\\\/]+$")
_AZ_CASE_RE = re.compile(r"\b\d+\s*[A-ZÄÖÜ]\s*\d+/\d+\b")


def _is_meaningful_token(s: str) -> bool:
    if not s:
        return False
    s = s.strip()
    if not s or s.lower() == "nan":
        return False
    # pure separators / borders → noise
    if _NOISE_RE.fullmatch(s):
        return False
    # a single stray character is likely noise
    if len(s) == 1 and not s.isalnum():
        return False
    return True


# Moved to class method


def _translate_numeric_token(token: str) -> str:
    if not token:
        return token
    if re.search(r"[0-9]", token):
        return token.translate(AMOUNT_TOKEN_TRANSLATE)
    if re.search(r"(EUR|€)", token, re.IGNORECASE):
        return token.translate(AMOUNT_TOKEN_TRANSLATE)
    return token


def normalize_line(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\u0080", "€")
    parts = re.split(r"(\s+)", text)
    pieces = []
    for part in parts:
        if not part:
            continue
        if part.isspace():
            pieces.append(part)
        else:
            pieces.append(_translate_numeric_token(part))
    joined = "".join(pieces)
    joined = re.sub(r"\s+", " ", joined).strip()
    joined = re.sub(r"\beur\b", "EUR", joined, flags=re.IGNORECASE)
    return joined


TOKEN_MATCH_TRANSLATE = str.maketrans(
    {
        "0": "o",
        "1": "l",
        "5": "s",
        "7": "t",
        "8": "b",
        "9": "g",
    }
)


def normalize_for_token_match(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text.translate(TOKEN_MATCH_TRANSLATE)


GG_LABEL_TRANSLATE = str.maketrans({"6": "G", "0": "G", "O": "G", "Q": "G", "C": "G", "€": "G"})
GG_EXTENDED_SUFFIX_RE = re.compile(
    r"^(?:GEMAESS|GEMAE?S|GEM|GEMAES)[A-Z0-9]*URT[A-Z0-9]*$"
)


def normalize_gg_candidate(text: str) -> str:
    if not text:
        return ""
    normalized = normalize_line(text)
    normalized = normalized.replace(":", " ")
    normalized = re.sub(r"[^A-Z0-9]", "", normalized.upper())
    translated = normalized.translate(GG_LABEL_TRANSLATE)
    if len(translated) >= 2:
        gg_pos = translated.find("GG")
        if gg_pos > 0:
            translated = translated[gg_pos:]
    return translated


def is_gg_label(text: str) -> bool:
    normalized = normalize_gg_candidate(text)
    if not normalized:
        return False
    if normalized == "GG":
        return True
    if normalized.startswith("GG"):
        remainder = normalized[2:]
        if not remainder:
            return True
        if GG_EXTENDED_SUFFIX_RE.match(remainder):
            return True
    return False


AMOUNT_RE = re.compile(
    r"(?:€\s*)?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d{2}|,-)?(?:\s*(?:EUR|€))?",
    re.IGNORECASE,
)
AMOUNT_CANDIDATE_RE = re.compile(
    r"(?:€\s*)?(?:\d{1,3}(?:[.,\s]\d{3})+|\d+)(?:[,\.]\d{2}|,-)?(?:\s*(?:EUR|€))?",
    re.IGNORECASE,
)
DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
INVOICE_RE = re.compile(r"\b\d{6,}\b")


def extract_amount_from_text(text: str, min_value=None):
    candidates = find_amount_candidates(text)
    if not candidates:
        return None

    min_decimal = None
    if min_value is not None:
        try:
            min_decimal = Decimal(str(min_value))
        except Exception:
            min_decimal = None

    if min_decimal is not None:
        candidates = [
            c
            for c in candidates
            if c.get("value") is not None and c["value"] >= min_decimal
        ]
        if not candidates:
            return None

    best = max(candidates, key=lambda c: c["value"])
    return best["display"] if best else None


def clean_amount_display(amount: str) -> str:
    if not amount:
        return amount
    amt = amount.strip()
    # ensure consistent spacing before currency suffix
    amt = re.sub(r"(\d)(EUR|€)$", r"\1 \2", amt, flags=re.IGNORECASE)
    amt = re.sub(r"\s+(EUR|€)$", r" \1", amt, flags=re.IGNORECASE)

    # remove leading zero-padded fragments that belong to invoice prefixes
    leading = re.match(r"^0\d{2}\s+", amt)
    if leading:
        remainder = amt[leading.end() :].strip()
        remainder = re.sub(r"(\d)(EUR|€)$", r"\1 \2", remainder, flags=re.IGNORECASE)
        remainder = re.sub(r"\s+(EUR|€)$", r" \1", remainder, flags=re.IGNORECASE)
        if remainder and AMOUNT_RE.fullmatch(remainder):
            amt = remainder
    return amt


def normalize_amount_candidate(raw_amount: str):
    if not raw_amount:
        return None

    text = unicodedata.normalize("NFKC", str(raw_amount))
    currency_match = re.search(r"(EUR|€)", text, re.IGNORECASE)
    currency_suffix = ""
    if currency_match:
        symbol = currency_match.group(1)
        currency_suffix = " EUR" if symbol.upper().startswith("EUR") else " €"

    text = re.sub(r"(EUR|€)", "", text, flags=re.IGNORECASE)
    text = text.replace("\u202f", " ").replace("\xa0", " ")
    text = text.replace("−", "-")
    text = text.replace("'", "").replace("`", "").replace("´", "")
    text = text.strip()
    text = re.sub(r",-+$", ",00", text)

    negative = False
    if text.startswith("-"):
        negative = True
        text = text[1:]

    text = text.replace(" ", "")
    text = re.sub(r"[^0-9,.-]", "", text)
    if not text:
        return None

    has_comma = "," in text
    has_dot = "." in text
    decimal_sep = None

    if has_comma and has_dot:
        decimal_sep = "," if text.rfind(",") > text.rfind(".") else "."
    elif has_comma:
        digits_after = len(re.sub(r"[^0-9]", "", text[text.rfind(",") + 1 :]))
        if 0 < digits_after <= 2:
            decimal_sep = ","
    elif has_dot:
        digits_after = len(re.sub(r"[^0-9]", "", text[text.rfind(".") + 1 :]))
        if 0 < digits_after <= 2:
            decimal_sep = "."

    if decimal_sep:
        sep_idx = text.rfind(decimal_sep)
        integer_raw = text[:sep_idx]
        decimal_raw = text[sep_idx + 1 :]
    else:
        integer_raw = text
        decimal_raw = ""

    integer_part = re.sub(r"[^0-9]", "", integer_raw)
    decimal_part = re.sub(r"[^0-9]", "", decimal_raw)

    if not integer_part:
        integer_part = "0"

    if not decimal_part:
        decimal_part = "00"
    elif len(decimal_part) == 1:
        decimal_part = f"{decimal_part}0"
    elif len(decimal_part) > 2:
        decimal_part = decimal_part[:2]

    try:
        value = Decimal(f"{int(integer_part)}.{decimal_part}")
    except (InvalidOperation, ValueError):
        return None

    if negative:
        value = -value

    formatted_int = f"{int(integer_part):,}".replace(",", ".")
    formatted = f"{formatted_int},{decimal_part}"
    if negative:
        formatted = f"-{formatted}"
    if currency_suffix:
        formatted = f"{formatted}{currency_suffix}"

    return clean_amount_display(formatted), value


def _amount_search_variants(normalized: str):
    variants = {normalized}
    if not normalized:
        return variants
    # keep separators tight like "1, 23" -> "1,23"
    compact_decimal = re.sub(r"([.,])\s+(?=\d)", r"\1", normalized)
    variants.add(compact_decimal)
    # ensure exactly one space before the currency suffix
    variants.add(re.sub(r"\s+(?=(?:EUR|€)\b)", " ", compact_decimal))
    return {v for v in variants if v}


def find_amount_candidates(text: str):
    if not text:
        return []
    normalized = normalize_line(text)

    # NEW: neutralize dates so they can't bleed into amount matches
    safe = DATE_RE.sub(" ", normalized)
    # Also neutralize invoice numbers to prevent them from sticking to amounts
    safe = re.sub(INVOICE_RE, " ", safe)

    seen = set()
    candidates = []
    for variant in _amount_search_variants(safe):
        for match in AMOUNT_CANDIDATE_RE.finditer(variant):
            raw = match.group(0)
            if (
                not re.search(r"(EUR|€)", raw, re.IGNORECASE)
                and "," not in raw
                and "." not in raw
            ):
                continue
            parsed = normalize_amount_candidate(raw)
            if not parsed:
                continue
            display, value = parsed
            key = (display, value)
            if key in seen:
                continue
            seen.add(key)
            candidates.append({"display": display, "value": value})
    return candidates


def build_streitwert_keywords(term):
    seen_keywords = set()
    keyword_candidates = []
    for candidate in [
        term,
        "Streitwert",
        "Streitwertes",
        "Streitwerts",
        "Streitgegenstand",
        "Streitgegenstandes",
        "Streitwert des Verfahrens",
        "Der Streitwert des Verfahrens",
        "Der Streitwert des Verfahrens wird",
        "Der Streitwert des Verfahrens wird auf",
        "Der Streitwert des Verfahrens wird auf bis zu",
        "Streitwert wurde",
        "Streitwert wird",
        "Streitwert wird auf",
        "Streitwert wird auf bis zu",
        "Streitwert wird bis",
        "Streitwert wird bis zu",
        "Der Streitwert wird auf",
        "Der Streitwert wird auf bis zu",
        "Der Streitwert wird bis",
        "Der Streitwert wird bis zu",
        "Die Streitwertfestsetzung",
        "Die Streitwertfestsetzung hatte",
        "Die Streitwertfestsetzung hatte einheitlich",
        "Die Streitwertfestsetzung hatte einheitlich auf",
        "Die Streitwertfestsetzung hatte einheitlich auf bis zu",
        "Streitwertfestsetzung",
        "Streitwertfestsetzung hatte",
        "Streitwertfestsetzung hatte einheitlich",
        "Streitwertfestsetzung hatte einheitlich auf",
        "Streitwertfestsetzung hatte einheitlich auf bis zu",
        "Streitwert beträgt",
        "Streitwert bis",
        "Streitwert bis Euro",
        "Streitwert bis EUR",
        "Streitwert bis zu",
        "Streitwert bis zu EUR",
        "Streitwert bis zu Euro",
        "wird auf",
        "wird vorläufig",
        "wird vorläufig auf",
        "wird vorlaufig",
        "wird vorlaufig auf",
        "der wird auf",
        "der wird vorläufig",
        "der wird vorläufig auf",
        "der wird vorlaufig",
        "der wird vorlaufig auf",
        "wird auf bis",
        "wird auf bis zu",
        "wird bis",
        "wird bis zu",
        "der wird bis",
        "der wird bis zu",
        "festgesetzt",
        "festgesetzt auf",
        "bis zu",
        "biszu",
        "bis euro",
        "gesetzt",
        "beträgt",
    ]:
        if candidate is None:
            continue
        key = str(candidate).strip()
        if not key:
            continue
        low = key.lower()
        if low in seen_keywords:
            continue
        seen_keywords.add(low)
        keyword_candidates.append(key)
    return keyword_candidates


DOC_LOADING_PATTERNS = (
    "dokumente werden geladen",
    "dokumente werden gel",
    "suche läuft",
    "suche lauft",
    "suche laeuft",
    "suche lauf",
    "daten werden geladen",
    "daten werden gel",
    "datei wird geladen",
    "datei wird gel",
    "bitte warten",
    "bitte warte",
    "wird geladen",
    "wird gelad",
    "wird geoffnet",
    "wird geöffnet",
    "werden vorbereitet",
    "wird vorbereitet",
    "lade daten",
    "lade datei",
    "laden",
)

LOG_SECTION_RE = re.compile(r"^\s*\[[^\]]*\]\s*Section:\s*(.+)$", re.IGNORECASE)
LOG_ENTRY_RE = re.compile(r"^\s*(\d{3}):\s*\(([^)]*)\)\s*->\s*(.*)$")
LOG_SOFT_RE = re.compile(r"^\s*soft:\s*(.*)$", re.IGNORECASE)
LOG_NORM_RE = re.compile(r"^\s*norm:\s*(.*)$", re.IGNORECASE)
LOG_KEYWORD_RE = re.compile(r"^\s*Keywords:\s*(.*)$", re.IGNORECASE)


def lines_from_tsv(tsv_df, scale=1):
    """
    From pytesseract data -> [(x,y,w,h,text), ...] top-to-bottom, left-to-right.
    Coordinates are normalised by the supplied scale factor (if any).
    """
    if tsv_df is None or tsv_df.empty:
        return []
    df = tsv_df.dropna(subset=["text"])
    df = df[df["conf"] > -1]
    try:
        scale_val = float(scale)
    except Exception:
        scale_val = 1.0
    scale_val = max(scale_val, 1.0)
    lines = []
    for (_, _, _), grp in df.groupby(["block_num", "par_num", "line_num"]):
        lefts = grp["left"]
        rights = grp["left"] + grp["width"]
        tops = grp["top"]
        bottoms = grp["top"] + grp["height"]
        xs = min(lefts) / scale_val
        ys = min(tops) / scale_val
        w = (max(rights) - min(lefts)) / scale_val
        h = (max(bottoms) - min(tops)) / scale_val
        txt = " ".join(str(t) for t in grp["text"] if str(t).strip())
        if txt.strip():
            lines.append(
                (
                    int(round(xs)),
                    int(round(ys)),
                    int(round(w)),
                    int(round(h)),
                    txt.strip(),
                )
            )
    lines.sort(key=lambda x: (x[1], x[0]))
    return lines


def _grab_region_color_generic(current_rect, rel_box, upscale):
    rx, ry, rw, rh = rel_to_abs(current_rect, rel_box)
    img = grab_xywh(rx, ry, rw, rh)
    try:
        scale_val = int(float(upscale))
    except Exception:
        scale_val = 3
    scale = max(1, scale_val)
    return upscale_pil(img, scale=scale), scale


# ---------- Normalization / parsing ----------
def normalize_line_soft(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\u0080", "€")
    parts = re.split(r"(\s+)", text)
    pieces = []
    for part in parts:
        if not part:
            continue
        if part.isspace():
            pieces.append(part)
            continue
        if re.search(r"[0-9]", part) or re.search(r"(EUR|€)", part, re.IGNORECASE):
            pieces.append(_translate_numeric_token(part))
        else:
            pieces.append(part.lower())
    joined = "".join(pieces)
    joined = re.sub(r"\s+", " ", joined).strip()
    joined = re.sub(r"\beur\b", "EUR", joined, flags=re.IGNORECASE)
    joined = re.sub(r"(\d+)\.(\d{2})\b", r"\1,\2", joined)
    return joined


def extract_amount_from_lines(lines, keyword=None, min_value=None):
    if not lines:
        return None, None

    processed = []
    for entry in lines:
        if isinstance(entry, (list, tuple)) and len(entry) == 5:
            x, y, w, h, text = entry
        else:
            y, text = entry
            x = w = h = None
        processed.append(
            {
                "y": y,
                "text": text or "",
                "norm": normalize_line_soft(text or ""),
                "candidates": find_amount_candidates(text or ""),
            }
        )

    combo_cache = {}

    def combo_info(idx, span):
        key = (idx, span)
        if key in combo_cache:
            return combo_cache[key]
        parts_text = []
        for offset in range(span):
            j = idx + offset
            if j >= len(processed):
                combo_cache[key] = ("", [])
                return combo_cache[key]
            parts_text.append(processed[j]["text"])
        combined_text = " ".join(part for part in parts_text if part).strip()
        combined_norm = normalize_line_soft(combined_text) if combined_text else ""
        if not combined_norm:
            combo_cache[key] = ("", [])
            return combo_cache[key]
        combo_cache[key] = (
            combined_norm,
            find_amount_candidates(combined_text),
        )
        return combo_cache[key]

    def candidate_variants(idx):
        if idx < 0 or idx >= len(processed):
            return
        info = processed[idx]
        seen = set()
        for cand in info["candidates"]:
            key = (cand["display"], cand["value"])
            if key in seen:
                continue
            seen.add(key)
            yield info["norm"], cand
        for span in (2, 3):
            combined_norm, combo_candidates = combo_info(idx, span)
            if not combo_candidates:
                continue
            for cand in combo_candidates:
                key = (cand["display"], cand["value"])
                if key in seen:
                    continue
                seen.add(key)
                yield combined_norm, cand

    try:
        min_decimal = Decimal(str(min_value)) if min_value is not None else None
    except Exception:
        min_decimal = None

    def pick_best(indices, required_keywords=None):
        best = None
        best_line = None
        best_value = None
        best_score = None
        if isinstance(required_keywords, str):
            required_terms = [required_keywords.strip().lower()]
        else:
            required_terms = [
                str(term).strip().lower()
                for term in (required_keywords or [])
                if str(term).strip()
            ]

        keyword_cache = {}

        def keyword_info(line_norm):
            cached = keyword_cache.get(line_norm)
            if cached is not None:
                return cached
            compact = re.sub(r"\s+", "", line_norm)
            alnum = re.sub(r"[^0-9a-z€]+", "", line_norm)
            norm_positions = []
            compact_positions = []
            alnum_positions = []
            for term in required_terms:
                if not term:
                    continue
                idx_norm = line_norm.find(term)
                if idx_norm != -1:
                    norm_positions.append(idx_norm)
                compact_term = re.sub(r"\s+", "", term)
                if compact_term:
                    idx_compact = compact.find(compact_term)
                    if idx_compact != -1:
                        compact_positions.append(idx_compact)
                alnum_term = re.sub(r"[^0-9a-z€]+", "", term)
                if alnum_term:
                    idx_alnum = alnum.find(alnum_term)
                    if idx_alnum != -1:
                        alnum_positions.append(idx_alnum)
            info = {
                "compact": compact,
                "alnum": alnum,
                "norm_positions": norm_positions,
                "compact_positions": compact_positions,
                "alnum_positions": alnum_positions,
            }
            keyword_cache[line_norm] = info
            return info

        for idx in indices:
            for line_text, candidate in candidate_variants(idx) or []:
                line_norm = (line_text or "").lower()
                if required_terms:
                    info_kw = keyword_info(line_norm)
                    compact_line = info_kw["compact"]
                    alnum_line = info_kw["alnum"]
                    norm_positions = info_kw["norm_positions"]
                    compact_positions = info_kw["compact_positions"]
                    alnum_positions = info_kw["alnum_positions"]
                    has_keyword = bool(
                        norm_positions or compact_positions or alnum_positions
                    )
                    if not has_keyword:
                        continue
                    priority = 1
                else:
                    compact_line = re.sub(r"\s+", "", line_norm)
                    alnum_line = re.sub(r"[^0-9a-z€]+", "", line_norm)
                    norm_positions = compact_positions = alnum_positions = []
                    priority = 0
                value = candidate.get("value")
                if value is None:
                    continue
                if min_decimal is not None and value < min_decimal:
                    continue
                cand_display = candidate.get("display", "")
                cand_norm = normalize_line_soft(cand_display).lower()
                cand_compact = re.sub(r"\s+", "", cand_norm)
                cand_alnum = re.sub(r"[^0-9a-z€]+", "", cand_norm)
                cand_idx = -1
                variant_used = "norm"
                if cand_norm:
                    cand_idx = line_norm.find(cand_norm)
                if cand_idx == -1 and cand_compact:
                    cand_idx = compact_line.find(cand_compact)
                    if cand_idx != -1:
                        variant_used = "compact"
                if cand_idx == -1 and cand_alnum:
                    cand_idx = alnum_line.find(cand_alnum)
                    if cand_idx != -1:
                        variant_used = "alnum"
                distance_score = 0
                after_keyword = 0
                if required_terms:
                    if (
                        not (norm_positions or compact_positions or alnum_positions)
                        or cand_idx == -1
                    ):
                        continue
                    if variant_used == "norm":
                        positions = (
                            norm_positions or compact_positions or alnum_positions
                        )
                    elif variant_used == "compact":
                        positions = (
                            compact_positions or norm_positions or alnum_positions
                        )
                    else:
                        positions = (
                            alnum_positions or compact_positions or norm_positions
                        )
                    diffs = [cand_idx - pos for pos in positions if cand_idx >= pos]
                    if diffs:
                        after_keyword = 1
                        distance_score = -min(diffs)
                    else:
                        continue
                position_score = -cand_idx if cand_idx >= 0 else float("-inf")
                score_tuple = (
                    priority + after_keyword,
                    after_keyword,
                    distance_score,
                    position_score,
                    value,
                )
                if best_score is None or score_tuple > best_score:
                    best = candidate
                    best_line = line_text
                    best_value = value
                    best_score = score_tuple
        if best:
            if min_decimal is not None and (
                best_value is None or best_value < min_decimal
            ):
                return None, None, None
            return best.get("display"), best_line, best_value
        return None, None, None

    if isinstance(keyword, (list, tuple, set)):
        keys = [str(k) for k in keyword if k is not None and str(k).strip()]
    elif keyword:
        keys = [str(keyword)]
    else:
        keys = []

    keys_norm = [
        normalize_line_soft(str(k).strip()).lower() for k in keys if str(k).strip()
    ]

    k_amt = k_line = None
    k_val = None

    def collect_candidate_indices(key_norm):
        keyword_indices = [
            idx
            for idx, info in enumerate(processed)
            if key_norm in info["norm"].lower()
        ]
        if not keyword_indices:
            return []
        offsets = [0, 1, -1, 2, -2]
        candidate_indices = []
        for idx in keyword_indices:
            for offset in offsets:
                candidate_indices.append(idx + offset)
        seen = set()
        ordered = []
        for idx in candidate_indices:
            if idx not in seen:
                seen.add(idx)
                ordered.append(idx)
        return ordered

    for key in keys:
        key_norm = normalize_line_soft(str(key).strip()).lower()
        if not key_norm:
            continue
        candidate_indices = collect_candidate_indices(key_norm)
        if not candidate_indices:
            continue
        amt, line, val = pick_best(candidate_indices, required_keywords=keys_norm)
        if not amt:
            continue
        current_val = val if val is not None else Decimal("0")
        stored_val = k_val if k_val is not None else Decimal("0")
        if k_amt is None or current_val > stored_val:
            k_amt, k_line, k_val = amt, line, val

    if k_amt:
        return k_amt, k_line

    if keys:
        return None, None

    g_amt, g_line, _ = pick_best(range(len(processed)))
    if g_amt:
        return g_amt, g_line

    combined_text = "\n".join(info["norm"] for info in processed if info["norm"]) or ""
    fallback = extract_amount_from_text(combined_text, min_value=min_decimal)
    if fallback:
        return fallback, combined_text

    return None, None


# ------------------ App Class ------------------
class RDPApp(tk.Tk):
    # --- Regex and constants (class-level) ---
    _KFB_RE = re.compile(
        r"(?<![0-9A-Za-z])k\s*[-./]?\s*f\s*[-./]?\s*b",
        re.IGNORECASE,
    )
    _KFB_WORD_RE = re.compile(
        r"kosten\s*festsetzungs\s*beschl(?:uss|uß|\.)?",
        re.IGNORECASE,
    )
    # European-style numbers like 1.234,56 or 1234,56 or 1 234,56
    _AMT_NUM_RE = re.compile(r"\b\d{1,3}(?:[\.\s]\d{3})*(?:,\d{2})?\b")
    # Hints that amount-in-words is present on page
    _WORDS_HINT_RE = re.compile(
        r"\b(?:euro|eur|tausend|hundert|einhundert|zweihundert|dreihundert|vierhundert|fuenf|fünf|sechs|sieben|acht|neun|zehn|elf|zwoelf|zwölf|zwanzig|dreissig|dreißig|vierzig|fuenfzig|fünfzig|sechzig|siebzig|achtzig|neunzig)\b",
        re.IGNORECASE,
    )

    # --- Doclist OCR + Click helpers ---

    def _ocr_doclist_rows(self):
        """
        OCR the calibrated doclist_region and return a list of row TEXTS (strings).
        Uses Tesseract's line grouping when available; otherwise groups by Y.
        """
        if not self._has("doclist_region"):
            self.log_print("[Doclist OCR] doclist_region not configured.")
            return []

        # Convert relative doclist region to absolute screen coords
        x, y, w, h = rel_to_abs(self.current_rect, self._get("doclist_region"))
        if w <= 0 or h <= 0:
            self.log_print(
                "[Doclist OCR] doclist_region has zero size after conversion."
            )
            return []
        img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
        df = do_ocr_data(img, lang=(self.lang_var.get().strip() or "deu+eng"), psm=6)

        if df is None or "text" not in df.columns:
            return []

        # Keep only meaningful tokens
        def _ok(s):
            return bool(s) and str(s).strip().lower() not in ("", "nan")

        df = df.copy()
        if "text" not in df.columns:  # safety
            return []
        df["text"] = df["text"].astype(str)
        df = df[df["text"].apply(_ok)]
        if df.empty:
            return []

        rows = []
        if {
            "block_num",
            "par_num",
            "line_num",
            "left",
            "top",
            "width",
            "height",
        }.issubset(df.columns):
            # Group by Tesseract line identifiers
            for (b, p, l), g in df.groupby(
                ["block_num", "par_num", "line_num"], sort=True
            ):
                g = g.sort_values("left")
                txt = " ".join(t.strip() for t in g["text"].tolist() if t.strip())
                if txt:
                    rows.append(txt)
        else:
            # Fallback: group by Y proximity
            df = df.sort_values("top")
            tol = max(8, int(img.height * 0.01))
            current_y = None
            buf = []
            for _, r in df.iterrows():
                if current_y is None or abs(int(r["top"]) - current_y) <= tol:
                    buf.append(str(r["text"]).strip())
                    if current_y is None:
                        current_y = int(r["top"])
                else:
                    line = " ".join(t for t in buf if t)
                    if line:
                        rows.append(line)
                    buf = [str(r["text"]).strip()]
                    current_y = int(r["top"])
            if buf:
                line = " ".join(t for t in buf if t)
                if line:
                    rows.append(line)

        # Light cleanup & de-dup short artifacts
        cleaned = []
        for s in rows:
            s2 = " ".join(s.split())
            if len(s2) >= 2:
                if not cleaned or cleaned[-1] != s2:
                    cleaned.append(s2)
        self.log_print(f"[Doclist OCR] lines: {len(cleaned)}")
        return cleaned

    def _ocr_doclist_rows_boxes(self):
        """
        OCR the calibrated doclist_region and return [(text, (lx,ty,rx,by)), ...]
        The box is in doclist image coords; map to screen via doclist_region.
        """
        if not self._has("doclist_region"):
            self.log_print("[Doclist OCR] doclist_region not configured.")
            return []

        # Convert relative doclist region to absolute screen coords
        X, Y, W, H = rel_to_abs(self.current_rect, self._get("doclist_region"))
        if W <= 0 or H <= 0:
            self.log_print(
                "[Doclist OCR] doclist_region has zero size after conversion."
            )
            return []
        img = self._grab_region_color(X, Y, W, H, upscale_x=self.upscale_var.get())
        df = do_ocr_data(img, lang=(self.lang_var.get().strip() or "deu+eng"), psm=6)
        if df is None or "text" not in df.columns:
            return []

        def _ok(s):
            return bool(s) and str(s).strip().lower() not in ("", "nan")

        df = df.copy()
        if "text" not in df.columns:  # safety
            return []
        df["text"] = df["text"].astype(str)
        df = df[df["text"].apply(_ok)]
        if df.empty:
            return []

        lines = []
        # Prefer Tesseract's line grouping when available
        if {
            "block_num",
            "par_num",
            "line_num",
            "left",
            "top",
            "width",
            "height",
        }.issubset(df.columns):
            for (_, _, line_num), g in df.groupby(
                ["block_num", "par_num", "line_num"], sort=True
            ):
                g = g.sort_values("left")
                txt = " ".join(t.strip() for t in g["text"].tolist() if t.strip())
                if not txt:
                    continue
                lx = int(g["left"].min())
                ty = int(g["top"].min())
                rx = int((g["left"] + g["width"]).max())
                by = int((g["top"] + g["height"]).max())
                lines.append((txt, (lx, ty, rx, by)))
        else:
            # Fallback: group by Y proximity
            df = df.sort_values("top")
            tol = max(8, int(img.height * 0.01))
            cur, cur_top = [], None

            def _flush():
                nonlocal lines, cur
                if not cur:
                    return
                lx = min(int(r["left"]) for _, r in cur)
                ty = min(int(r["top"]) for _, r in cur)
                rx = max(int(r["left"] + r["width"]) for _, r in cur)
                by = max(int(r["top"] + r["height"]) for _, r in cur)
                txt = " ".join(
                    str(r["text"]).strip()
                    for _, r in sorted(cur, key=lambda t: int(t[1]["left"]))
                    if str(r["text"]).strip()
                )
                if txt:
                    lines.append((txt, (lx, ty, rx, by)))
                cur = []

            for it in df.iterrows():
                _, r = it
                y = int(r["top"])
                if cur_top is None or abs(y - cur_top) <= tol:
                    cur.append(it)
                    cur_top = y if cur_top is None else cur_top
                else:
                    _flush()
                    cur = [it]
                    cur_top = y
            _flush()

        self.log_print(f"[Doclist OCR] lines+boxes: {len(lines)}")
        return lines

    def _click_doclist_row(self, row_idx: int):
        """
        Click the center of the given row index inside the doclist_region using OCR boxes.
        Returns True on success, False otherwise.
        """
        if row_idx is None or row_idx < 0:
            return False
        rows = self._ocr_doclist_rows_boxes()
        if not rows or row_idx >= len(rows):
            return False

        # target box in doclist image coords (relative to the captured doclist image)
        _, (lx, ty, rx, by) = rows[row_idx]
        cx = (lx + rx) // 2
        cy = (ty + by) // 2

        # Map to absolute screen coords using calibrated doclist_region (relative -> absolute)
        try:
            rx, ry, rw, rh = rel_to_abs(self.current_rect, self._get("doclist_region"))
            abs_x = int(rx + cx)
            abs_y = int(ry + cy)
        except Exception:
            return False

        try:
            # Ensure RDP window focused before clicking
            try:
                Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
            except Exception:
                pass
            pyautogui.click(abs_x, abs_y)
            time.sleep(0.08)
            return True
        except Exception as e:
            self.log_print(f"[Doclist OCR] click failed: {e}")
            return False

    # --- Instance table helpers ---

    def pick_instance_region(self):
        """Two-click calibration for the AZ Instanz table region (headers + a bit below)."""
        rb = self._two_click_box(
            "Hover TOP-LEFT of the AZ Instanz table (include headers), then press OK.",
            "Hover BOTTOM-RIGHT (include a little area below headers), then press OK.",
        )
        if rb:
            self.cfg["instance_region"] = rb
            try:
                # if you created a readonly text field for previewing values, update it
                self.instance_var.set(
                    f"{rb[0]:.3f}, {rb[1]:.3f}, {rb[2]:.3f}, {rb[3]:.3f}"
                )
            except Exception:
                pass
            self.log_print(f"Instance table region set: {rb}")

    def _grab_instance_region(self):
        """Grab the calibrated instance region as a PIL image."""
        if not getattr(self, "current_rect", None):
            self.connect_rdp()
            if not self.current_rect:
                self.log_print(
                    "No active RDP rectangle; cannot capture instance region."
                )
                return None, None

        if not self._has("instance_region"):
            self.log_print(
                "Instance region not configured. Use 'Pick Instance Table Region'."
            )
            return None, None

        # Convert relative coordinates to absolute screen coordinates
        x, y, w, h = rel_to_abs(self.current_rect, self._get("instance_region"))

        # Grab the region using our class method
        img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
        scale = max(1, int(self.upscale_var.get() or 3))

        return img, scale

    def extract_instance_columns_text(self):
        """
        Split the calibrated region into 3 equal columns (I/II/III),
        OCR BELOW the header row, return raw text per column (STRICT).
        """
        img, _ = self._grab_instance_region()
        if img is None:
            return None

        w, h = img.width, img.height
        # start scanning below headers (feel free to set 0.50 in config if header is tall)
        rel_top = float(self.cfg.get("instance_row_rel_top", 0.45))
        y0 = max(0, min(int(h * rel_top), h - 1))
        y1 = max(y0 + 1, h)

        thirds = [(0, w // 3), (w // 3, 2 * w // 3), (2 * w // 3, w)]
        col_texts = []
        for x0, x1 in thirds:
            crop = self._safe_crop(img, (x0, y0, x1, y1))
            col_texts.append(self._ocr_text_strict(crop))  # <— strict OCR here

        return {
            "preview_image": img,
            "below_header_y": y0,
            "col_I_text": col_texts[0],
            "col_II_text": col_texts[1],
            "col_III_text": col_texts[2],
        }

    # --- Instance detection (strict) shared helpers ---

    _NOISE_RE = re.compile(r"^[_\-–—\|:;.,'\"`~^°()+\[\]{}<>\\\/]+$")
    _AZ_CASE_RE = re.compile(r"\b\d+\s*[A-ZÄÖÜ]\s*\d+/\d+\b")  # e.g., "15 O 9715/21"

    def _is_meaningful_token(self, s: str) -> bool:
        if not s:
            return False
        s = str(s).strip()
        if not s or s.lower() == "nan":
            return False
        if self._NOISE_RE.fullmatch(s):
            return False
        if len(s) == 1 and not s.isalnum():
            return False
        return True

    def _has_meaningful_content(self, s: str) -> bool:
        if not s:
            return False
        s2 = s.strip()
        if not s2:
            return False
        if self._AZ_CASE_RE.search(
            s2
        ):  # a real AZ pattern conclusively counts as present
            return True
        if not any(ch.isalnum() for ch in s2):  # ignore lone separators/graphics
            return False
        return len(s2.replace(" ", "")) >= 2  # avoid 1-char artifacts

    def _ocr_text_strict(self, pil_image):
        lang = (
            self.lang_var.get().strip() if hasattr(self, "lang_var") else ""
        ) or "deu+eng"
        df = do_ocr_data(pil_image, lang=lang, psm=6)
        try:
            raw = df["text"].tolist() if "text" in df.columns else []
        except Exception:
            raw = []
        tokens = []
        for t in raw:
            s = "" if t is None else str(t)
            if self._is_meaningful_token(s):
                tokens.append(s.strip())
        return " ".join(tokens).strip()

    # --- Instance detection (strict) shared helpers ---

    _NOISE_RE = re.compile(r"^[_\-–—\|:;.,'\"`~^°()+\[\]{}<>\\\/]+$")
    _AZ_CASE_RE = re.compile(r"\b\d+\s*[A-ZÄÖÜ]\s*\d+/\d+\b")  # e.g., "15 O 9715/21"

    def _is_meaningful_token(self, s: str) -> bool:
        if not s:
            return False
        s = str(s).strip()
        if not s or s.lower() == "nan":
            return False
        if self._NOISE_RE.fullmatch(s):
            return False
        if len(s) == 1 and not s.isalnum():
            return False
        return True

    def _has_meaningful_content(self, s: str) -> bool:
        if not s:
            return False
        s2 = s.strip()
        if not s2:
            return False
        if self._AZ_CASE_RE.search(
            s2
        ):  # a real AZ pattern conclusively counts as present
            return True
        if not any(ch.isalnum() for ch in s2):  # ignore lone separators/graphics
            return False
        return len(s2.replace(" ", "")) >= 2  # avoid 1-char artifacts

    def _ocr_text_strict(self, pil_image):
        lang = (
            self.lang_var.get().strip() if hasattr(self, "lang_var") else ""
        ) or "deu+eng"
        df = do_ocr_data(pil_image, lang=lang, psm=6)
        try:
            raw = df["text"].tolist() if "text" in df.columns else []
        except Exception:
            raw = []
        tokens = []
        for t in raw:
            s = "" if t is None else str(t)
            if self._is_meaningful_token(s):
                tokens.append(s.strip())
        return " ".join(tokens).strip()

    def _ocr_text(self, pil_image):
        """Run OCR on a PIL image and return meaningful text only (filters noise)."""
        return self._ocr_text_strict(pil_image)  # Use strict by default

    def extract_instance_columns_text(self):
        """
        Split the calibrated region into 3 equal columns (I/II/III),
        OCR BELOW the header row, return raw text per column (STRICT).
        """
        img, _ = self._grab_instance_region()
        if img is None:
            return None

        w, h = img.width, img.height
        # start scanning below headers (feel free to set 0.50 in config if header is tall)
        rel_top = float(self.cfg.get("instance_row_rel_top", 0.45))
        y0 = max(0, min(int(h * rel_top), h - 1))
        y1 = max(y0 + 1, h)

        thirds = [(0, w // 3), (w // 3, 2 * w // 3), (2 * w // 3, w)]
        col_texts = []
        for x0, x1 in thirds:
            crop = self._safe_crop(img, (x0, y0, x1, y1))
            col_texts.append(self._ocr_text_strict(crop))  # <— strict OCR here

        return {
            "preview_image": img,
            "below_header_y": y0,
            "col_I_text": col_texts[0],
            "col_II_text": col_texts[1],
            "col_III_text": col_texts[2],
        }

    def test_instance_extraction(self):
        """UI action wired to the button: preview + log the raw texts and detected instance."""
        from datetime import datetime
        import os

        try:
            if hasattr(self, "apply_paths_to_tesseract"):
                self.apply_paths_to_tesseract()

            info = self.detect_instance(prefix="(Test) ")
            if info is None:
                return

            # Optional extra preview: draw a line where scanning starts
            data = self.extract_instance_columns_text()
            img = data["preview_image"].copy()
            try:
                from PIL import ImageDraw

                y = data["below_header_y"]
                y = max(0, min(int(y), img.height - 1))
                ImageDraw.Draw(img).line([(0, y), (img.width, y)], width=2)
            except Exception:
                pass
            self.show_preview(img)

            # Small TXT dump (unchanged)
            log_dir = globals().get("LOG_DIR", "./log")
            os.makedirs(log_dir, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(log_dir, f"instance_ocr_{ts}.txt")
            with open(path, "w", encoding="utf-8") as f:
                f.write("Instance OCR (below headers)\n")
                f.write(f"StartY: {data['below_header_y']}\n")
                f.write(f"Inst1: {data['col_I_text']}\n")
                f.write(f"Inst2: {data['col_II_text']}\n")
                f.write(f"Inst3: {data['col_III_text']}\n")
                f.write(f"Detected: {info['instance']}\n")
            self.log_print(f"Saved instance OCR report → {path}")

        except Exception as e:
            try:
                messagebox.showerror("Test Instance Extraction", f"Failed: {e}")
            except Exception:
                self.log_print(f"Test Instance Extraction failed: {e}")

    def detect_instance(self, prefix=""):
        """
        Detects whether AZ I., AZ II., AZ III. Instanz have meaningful text below them.
        Uses strict OCR and content checking to filter out noise like borders/punctuation.
        """
        data = self.extract_instance_columns_text()
        if not data:
            return None

        I = self._has_meaningful_content(data["col_I_text"])
        II = self._has_meaningful_content(data["col_II_text"])
        III = self._has_meaningful_content(data["col_III_text"])

        if [I, II, III] == [True, False, False]:
            inst = 1
        elif [I, II, III] == [True, True, False]:
            inst = 2
        elif [I, II, III] == [True, True, True]:
            inst = 3
        else:
            inst = max((i + 1 for i, v in enumerate([I, II, III]) if v), default=None)

        # unified logging (the workflows will show the same as the test)
        self.log_print(f"{prefix}Instance columns (strict):")
        self.log_print(f"{prefix}  I  : {(data['col_I_text'] or '(empty)').strip()}")
        self.log_print(f"{prefix}  II : {(data['col_II_text'] or '(empty)').strip()}")
        self.log_print(f"{prefix}  III: {(data['col_III_text'] or '(empty)').strip()}")
        self.log_print(
            f"{prefix}Instance detection: I={I}, II={II}, III={III} → "
            f"{('undetermined' if inst is None else f'{inst}. Instanz')}"
        )

        return {"inst1": I, "inst2": II, "inst3": III, "instance": inst}

        # Compute row to start scanning (below the header labels)
        rel_top = float(self.cfg.get("instance_row_rel_top", 0.45))
        w, h = img.width, img.height
        y0 = max(0, min(int(h * rel_top), h - 1))
        y1 = max(y0 + 1, h)

        # 3 equal columns (I / II / III Instanz)
        thirds = [(0, w // 3), (w // 3, 2 * w // 3), (2 * w // 3, w)]

        present = []
        for i, (x0, x1) in enumerate(thirds, start=1):
            text = self._ocr_text_strict(self._safe_crop(img, (x0, y0, x1, y1)))
            has_text = self._has_meaningful_content(text)
            present.append(has_text)
            self.log_print(
                f"{prefix}Instanz {i} slice OCR: {'(text found)' if has_text else '(empty)'}"
            )

        # Decide the 'current' instance:
        # If I only -> 1st; I & II -> 2nd; I & II & III -> 3rd.
        # If something unexpected (e.g., II only), fall back to max index with text.
        instance_no = None
        if present == [True, False, False]:
            instance_no = 1
        elif present == [True, True, False]:
            instance_no = 2
        elif present == [True, True, True]:
            instance_no = 3
        else:
            # Fallback: choose highest active column index (1-based)
            if any(present):
                instance_no = max(i + 1 for i, v in enumerate(present) if v)

        # Log the summary
        I, II, III = present
        self.log_print(
            f"{prefix}Instance detection: I={I}, II={II}, III={III} -> "
            f"{('undetermined' if instance_no is None else f'{instance_no}. Instanz')}"
        )

        return {"inst1": I, "inst2": II, "inst3": III, "instance": instance_no}

    def __init__(self):
        # Initialize the Tk window first
        super().__init__()

        # Set window title and size
        self.title("RDP Automation (Tkinter)")
        self.geometry("1220x860")
        self.minsize(1080, 760)

        # Load configuration
        self.cfg = load_cfg()

        # Initialize instance variables
        self.current_rect = None
        self.ocr_preview_imgtk = None
        self._current_profile_sub_region = None
        self.capture_countdown_seconds = 3
        self.live_preview_window = None
        self.live_preview_label = None
        self.live_preview_imgtk = None
        self.live_preview_running = False
        self._ocr_log_paths = {}
        self._preview_ready_event = threading.Event()
        self._preview_last_token = 0
        self._preview_target_token = 0

        # Initialize MSS in the main thread
        get_mss()

        # Create all GUI widgets
        self.create_widgets()

        # Set up window close handler
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Call update to ensure window is fully initialized
        self.update()

    def _on_closing(self):
        """Clean up resources before closing."""
        try:
            # Clean up MSS resources
            if hasattr(_thread_local, "sct"):
                _thread_local.sct.close()
                delattr(_thread_local, "sct")

            # Clean up any preview windows
            if self.live_preview_window is not None:
                try:
                    self.live_preview_window.destroy()
                except:
                    pass

            # Stop live preview if running
            self.live_preview_running = False

            # Ensure main window gets destroyed properly
            self.quit()
            self.destroy()
        except:
            # Force quit if normal cleanup fails
            self.destroy()

    def _has(self, region_name: str) -> bool:
        """Check if a region is configured in the current configuration."""
        return (
            region_name in self.cfg
            and isinstance(self.cfg[region_name], list)
            and len(self.cfg[region_name]) == 4
            and all(isinstance(x, (int, float)) for x in self.cfg[region_name])
        )

    def _get(self, region_name: str) -> tuple[float, float, float, float]:
        """Get the configured coordinates for a region as (x, y, w, h)."""
        if not self._has(region_name):
            raise ValueError(f"Region {region_name} not configured")
        return tuple(self.cfg[region_name])

    def create_widgets(self):
        # --- Left frame with notebook sections ---
        left_container = ttk.Frame(self, padding=10)
        left_container.pack(side=tk.LEFT, fill=tk.Y)

        notebook = ttk.Notebook(left_container)
        notebook.pack(fill=tk.BOTH, expand=True)

        general_tab = ttk.Frame(notebook)
        calibration_tab = ttk.Frame(notebook)
        streit_tab = ttk.Frame(notebook)
        rechn_tab = ttk.Frame(notebook)
        fees_tab = ttk.Frame(notebook)
        log_tab = ttk.Frame(notebook)
        ocr_tab = ttk.Frame(notebook)

        notebook.add(general_tab, text="General")
        notebook.add(calibration_tab, text="Calibration")
        notebook.add(streit_tab, text="Streitwert")
        notebook.add(rechn_tab, text="Rechnungen")
        notebook.add(fees_tab, text="Fees")

        # Add Fees UI components
        fees_frame = ttk.LabelFrame(fees_tab, text="Calibration & Options")
        fees_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        # Calibration buttons
        fcal = ttk.Frame(fees_frame)
        fcal.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            fcal,
            text="Pick File Search Region",
            command=self.pick_fees_file_search_region,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            fcal,
            text="Pick Seiten Region",
            command=self.pick_fees_seiten_region,
        ).pack(side=tk.LEFT, padx=2)

        # Show the current coordinates
        ttk.Label(fees_frame, text="File Search Region (l%, t%, w%, h%)").pack(
            anchor="w", pady=(6, 0)
        )
        fs = self.cfg.get("fees_file_search_region", [0, 0, 0, 0])
        self.fees_file_search_var = tk.StringVar(
            value=f"{fs[0]:.3f}, {fs[1]:.3f}, {fs[2]:.3f}, {fs[3]:.3f}"
        )
        ttk.Entry(
            fees_frame,
            textvariable=self.fees_file_search_var,
            width=40,
            state="readonly",
        ).pack(anchor="w", pady=(0, 4))

        ttk.Label(fees_frame, text="Seiten Region (l%, t%, w%, h%)").pack(
            anchor="w", pady=(2, 0)
        )
        st = self.cfg.get("fees_seiten_region", [0, 0, 0, 0])
        self.fees_seiten_var = tk.StringVar(
            value=f"{st[0]:.3f}, {st[1]:.3f}, {st[2]:.3f}, {st[3]:.3f}"
        )
        ttk.Entry(
            fees_frame,
            textvariable=self.fees_seiten_var,
            width=40,
            state="readonly",
        ).pack(anchor="w", pady=(0, 4))

        # Configuration options
        ttk.Label(fees_frame, text="File search token (e.g., KFB)").pack(
            anchor="w", pady=(6, 0)
        )
        self.fees_search_var = tk.StringVar(
            value=self.cfg.get("fees_search_token", "KFB")
        )
        ttk.Entry(fees_frame, textvariable=self.fees_search_var, width=16).pack(
            anchor="w"
        )

        ttk.Label(fees_frame, text="Bad prefixes (semicolon-separated)").pack(
            anchor="w", pady=(6, 0)
        )
        self.fees_bad_var = tk.StringVar(
            value=self.cfg.get("fees_bad_prefixes", "SVRAGS;SVR-AGS;Skrags;SV RAGS")
        )
        bad_frame = ttk.Frame(fees_frame)
        bad_frame.pack(anchor="w")
        ttk.Entry(bad_frame, textvariable=self.fees_bad_var, width=40).pack(
            side=tk.LEFT
        )
        ttk.Button(
            bad_frame,
            text="Edit...",
            command=self.edit_fees_bad_prefixes,
        ).pack(side=tk.LEFT, padx=4)

        # Maximum page clicks
        mp = ttk.Frame(fees_frame)
        mp.pack(anchor="w", pady=(6, 0))
        ttk.Label(mp, text="Max page clicks").pack(side=tk.LEFT)
        self.fees_pages_max_var = tk.StringVar(
            value=str(self.cfg.get("fees_pages_max_clicks", 12))
        )
        ttk.Entry(mp, textvariable=self.fees_pages_max_var, width=6).pack(
            side=tk.LEFT, padx=6
        )

        # Options
        self.fees_skip_waits_var = tk.BooleanVar(
            value=self.cfg.get("fees_overlay_skip_waits", True)
        )
        ttk.Checkbutton(
            fees_frame,
            text="Only wait for loading overlays (ignore manual delays)",
            variable=self.fees_skip_waits_var,
        ).pack(anchor="w", pady=(6, 0))

        # Output file
        ttk.Label(fees_frame, text="Output CSV").pack(anchor="w", pady=(6, 0))
        self.fees_csv_var = tk.StringVar(
            value=self.cfg.get("fees_csv_path", "fees_results.csv")
        )
        ttk.Entry(fees_frame, textvariable=self.fees_csv_var, width=40).pack(
            anchor="w", pady=(0, 4)
        )

        # Action buttons
        ttk.Button(
            fees_frame,
            text="Test Fee Extraction",
            command=self.test_fees,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Button(
            fees_frame,
            text="Test Seiten Clicks",
            command=self.test_fees_seiten_clicks,
        ).pack(anchor="w", pady=(4, 0))

        ttk.Button(
            fees_frame,
            text="Run Fee Processing",
            command=self.run_fees,
        ).pack(anchor="w", pady=(4, 0))

        notebook.add(log_tab, text="Log")
        notebook.add(ocr_tab, text="OCR / Profiles")

        # --- General tab: RDP, Excel, timing ---
        rdp_frame = ttk.LabelFrame(general_tab, text="RDP Connection")
        rdp_frame.pack(fill=tk.X, pady=4)
        ttk.Label(rdp_frame, text="RDP Window Title (Regex)").pack(anchor="w")
        self.rdp_var = tk.StringVar(value=self.cfg["rdp_title_regex"])
        ttk.Entry(rdp_frame, textvariable=self.rdp_var, width=52).pack(
            anchor="w", pady=(0, 6)
        )

        excel_frame = ttk.LabelFrame(general_tab, text="Excel Source")
        excel_frame.pack(fill=tk.X, pady=4)
        ttk.Label(excel_frame, text="Excel Path").pack(anchor="w")
        xframe = ttk.Frame(excel_frame)
        xframe.pack(anchor="w", fill=tk.X)
        self.xls_var = tk.StringVar(value=self.cfg["excel_path"])
        ttk.Entry(xframe, textvariable=self.xls_var, width=42).pack(
            side=tk.LEFT, pady=2
        )
        ttk.Button(xframe, text="Browse", command=self.browse_excel).pack(
            side=tk.LEFT, padx=6
        )

        row1 = ttk.Frame(excel_frame)
        row1.pack(anchor="w", pady=(6, 0))
        ttk.Label(row1, text="Sheet").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar(value=str(self.cfg["excel_sheet"]))
        ttk.Entry(row1, textvariable=self.sheet_var, width=10).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Label(row1, text="Start cell (e.g., B2)").pack(side=tk.LEFT)
        self.start_cell_var = tk.StringVar(value=self.cfg.get("start_cell", ""))
        ttk.Entry(row1, textvariable=self.start_cell_var, width=10).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Label(row1, text="Max rows (0=all)").pack(side=tk.LEFT)
        self.max_rows_var = tk.StringVar(value=str(self.cfg.get("max_rows", 0)))
        ttk.Entry(row1, textvariable=self.max_rows_var, width=7).pack(
            side=tk.LEFT, padx=6
        )

        ttk.Label(
            excel_frame, text="(Optional) Input column name (if Start cell empty)"
        ).pack(anchor="w", pady=(6, 0))
        self.col_var = tk.StringVar(value=self.cfg["input_column"])
        ttk.Entry(excel_frame, textvariable=self.col_var, width=20).pack(
            anchor="w", pady=(0, 6)
        )

        ttk.Label(excel_frame, text="Results CSV").pack(anchor="w")
        self.csv_var = tk.StringVar(value=self.cfg["results_csv"])
        ttk.Entry(excel_frame, textvariable=self.csv_var, width=42).pack(
            anchor="w", pady=(0, 6)
        )

        tess_frame = ttk.LabelFrame(general_tab, text="Tesseract")
        tess_frame.pack(fill=tk.X, pady=4)
        ttk.Label(tess_frame, text="Tesseract Path (exe or folder)").pack(anchor="w")
        tframe = ttk.Frame(tess_frame)
        tframe.pack(anchor="w", fill=tk.X)
        self.tess_var = tk.StringVar(value=self.cfg["tesseract_path"])
        ttk.Entry(tframe, textvariable=self.tess_var, width=42).pack(
            side=tk.LEFT, pady=2
        )
        ttk.Button(tframe, text="Browse", command=self.browse_tesseract).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Label(tess_frame, text="OCR language (e.g., deu+eng)").pack(
            anchor="w", pady=(6, 0)
        )
        self.lang_var = tk.StringVar(value=self.cfg.get("tesseract_lang", "deu+eng"))
        ttk.Entry(tess_frame, textvariable=self.lang_var, width=16).pack(
            anchor="w", pady=(0, 6)
        )

        timing_frame = ttk.LabelFrame(general_tab, text="Timing & Typing")
        timing_frame.pack(fill=tk.X, pady=4)
        r1 = ttk.Frame(timing_frame)
        r1.pack(anchor="w")
        ttk.Label(r1, text="Typing delay (sec/char)").pack(side=tk.LEFT)
        self.type_var = tk.StringVar(value=str(self.cfg["type_delay"]))
        ttk.Entry(r1, textvariable=self.type_var, width=8).pack(side=tk.LEFT, padx=6)

        r2 = ttk.Frame(timing_frame)
        r2.pack(anchor="w", pady=(4, 0))
        ttk.Label(r2, text="Post-search wait (sec)").pack(side=tk.LEFT)
        self.wait_var = tk.StringVar(value=str(self.cfg["post_search_wait"]))
        ttk.Entry(r2, textvariable=self.wait_var, width=8).pack(side=tk.LEFT, padx=6)

        r3 = ttk.Frame(timing_frame)
        r3.pack(anchor="w", pady=(6, 0))
        ttk.Label(r3, text="Typing test text").pack(side=tk.LEFT)
        self.typing_test_var = tk.StringVar(
            value=self.cfg.get("typing_test_text", "TEST123")
        )
        ttk.Entry(r3, textvariable=self.typing_test_var, width=18).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(r3, text="Test Typing", command=self.test_typing).pack(
            side=tk.LEFT, padx=4
        )

        # --- Calibration tab ---
        cal_frame = ttk.LabelFrame(calibration_tab, text="RDP Calibration")
        cal_frame.pack(fill=tk.BOTH, expand=True, pady=4)
        cframe = ttk.Frame(cal_frame)
        cframe.pack(anchor="w", pady=(0, 4))
        ttk.Button(cframe, text="Connect RDP", command=self.connect_rdp).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            cframe, text="Pick Search Point", command=self.pick_search_point
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            cframe, text="Pick Result Region", command=self.pick_result_region
        ).pack(side=tk.LEFT, padx=2)

        ttk.Label(cal_frame, text="Search Point (x%, y%)").pack(anchor="w", pady=(4, 0))
        self.sp_var = tk.StringVar(
            value=f"{self.cfg['search_point'][0]:.3f}, {self.cfg['search_point'][1]:.3f}"
        )
        ttk.Entry(cal_frame, textvariable=self.sp_var, width=30).pack(anchor="w")

        ttk.Label(cal_frame, text="Result Region (l%, t%, w%, h%)").pack(
            anchor="w", pady=(6, 0)
        )
        rr = self.cfg["result_region"]
        self.rr_var = tk.StringVar(
            value=f"{rr[0]:.3f}, {rr[1]:.3f}, {rr[2]:.3f}, {rr[3]:.3f}"
        )
        ttk.Entry(cal_frame, textvariable=self.rr_var, width=40).pack(anchor="w")

        # Instance region calibration controls
        ttk.Button(
            cframe, text="Pick Instance Table Region", command=self.pick_instance_region
        ).pack(side=tk.LEFT, padx=2)

        ttk.Label(cal_frame, text="Instance Region (l%, t%, w%, h%)").pack(
            anchor="w", pady=(6, 0)
        )
        ir = self.cfg.get("instance_region", [0, 0, 0, 0])
        self.instance_var = tk.StringVar(
            value=f"{ir[0]:.3f}, {ir[1]:.3f}, {ir[2]:.3f}, {ir[3]:.3f}"
        )
        ttk.Entry(
            cal_frame, textvariable=self.instance_var, width=40, state="readonly"
        ).pack(anchor="w")

        # --- Streitwert tab ---
        streit_frame = ttk.LabelFrame(streit_tab, text="Calibration & Filters")
        streit_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        cal = ttk.Frame(streit_frame)
        cal.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal,
            text="Pick Doc List Region",
            command=self.pick_doclist_region,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            cal,
            text="Pick PDF Search Point",
            command=self.pick_pdf_search_point,
        ).pack(side=tk.LEFT, padx=2)

        cal2 = ttk.Frame(streit_frame)
        cal2.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal2,
            text="Pick PDF Results Button",
            command=self.pick_pdf_hits_point,
        ).pack(side=tk.LEFT, padx=2)
        hits_pt = self.cfg.get("pdf_hits_point")
        hits_txt = (
            f"{hits_pt[0]:.3f}, {hits_pt[1]:.3f}"
            if isinstance(hits_pt, (list, tuple)) and len(hits_pt) == 2
            else ""
        )
        self.pdf_hits_var = tk.StringVar(value=hits_txt)
        ttk.Entry(
            cal2,
            textvariable=self.pdf_hits_var,
            width=20,
            state="readonly",
        ).pack(side=tk.LEFT, padx=4)

        cal2b = ttk.Frame(streit_frame)
        cal2b.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal2b,
            text="Pick 2nd PDF Result Button",
            command=self.pick_pdf_second_hits_point,
        ).pack(side=tk.LEFT, padx=2)
        hits2_pt = self.cfg.get("pdf_hits_second_point")
        hits2_txt = (
            f"{hits2_pt[0]:.3f}, {hits2_pt[1]:.3f}"
            if isinstance(hits2_pt, (list, tuple)) and len(hits2_pt) == 2
            else ""
        )
        self.pdf_hits2_var = tk.StringVar(value=hits2_txt)
        ttk.Entry(
            cal2b,
            textvariable=self.pdf_hits2_var,
            width=20,
            state="readonly",
        ).pack(side=tk.LEFT, padx=4)

        cal2c = ttk.Frame(streit_frame)
        cal2c.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal2c,
            text="Pick 3rd PDF Result Button",
            command=self.pick_pdf_third_hits_point,
        ).pack(side=tk.LEFT, padx=2)
        hits3_pt = self.cfg.get("pdf_hits_third_point")
        hits3_txt = (
            f"{hits3_pt[0]:.3f}, {hits3_pt[1]:.3f}"
            if isinstance(hits3_pt, (list, tuple)) and len(hits3_pt) == 2
            else ""
        )
        self.pdf_hits3_var = tk.StringVar(value=hits3_txt)
        ttk.Entry(
            cal2c,
            textvariable=self.pdf_hits3_var,
            width=20,
            state="readonly",
        ).pack(side=tk.LEFT, padx=4)

        cal3 = ttk.Frame(streit_frame)
        cal3.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal3,
            text="Pick PDF Text Region",
            command=self.pick_pdf_text_region,
        ).pack(side=tk.LEFT, padx=2)

        cal4 = ttk.Frame(streit_frame)
        cal4.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal4,
            text="Pick View Button",
            command=self.pick_doc_view_point,
        ).pack(side=tk.LEFT, padx=2)
        view_pt = self.cfg.get("doc_view_point")
        view_txt = (
            f"{view_pt[0]:.3f}, {view_pt[1]:.3f}"
            if isinstance(view_pt, (list, tuple)) and len(view_pt) == 2
            else ""
        )
        self.doc_view_var = tk.StringVar(value=view_txt)
        ttk.Entry(
            cal4,
            textvariable=self.doc_view_var,
            width=20,
            state="readonly",
        ).pack(side=tk.LEFT, padx=4)

        cal5 = ttk.Frame(streit_frame)
        cal5.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            cal5,
            text="Pick PDF Close Button",
            command=self.pick_pdf_close_point,
        ).pack(side=tk.LEFT, padx=2)
        close_pt = self.cfg.get("pdf_close_point")
        close_txt = (
            f"{close_pt[0]:.3f}, {close_pt[1]:.3f}"
            if isinstance(close_pt, (list, tuple)) and len(close_pt) == 2
            else ""
        )
        self.pdf_close_var = tk.StringVar(value=close_txt)
        ttk.Entry(
            cal5,
            textvariable=self.pdf_close_var,
            width=20,
            state="readonly",
        ).pack(side=tk.LEFT, padx=4)

        ttk.Label(streit_frame, text="Include tokens (comma-separated)").pack(
            anchor="w", pady=(6, 0)
        )
        self.includes_var = tk.StringVar(value=self.cfg.get("includes", "Urt,SWB,SW"))
        ttk.Entry(streit_frame, textvariable=self.includes_var, width=40).pack(
            anchor="w"
        )

        ttk.Label(streit_frame, text="Exclude tokens (comma-separated)").pack(
            anchor="w", pady=(6, 0)
        )
        self.excludes_var = tk.StringVar(value=self.cfg.get("excludes", "SaM,KLE"))
        ttk.Entry(streit_frame, textvariable=self.excludes_var, width=40).pack(
            anchor="w"
        )

        self.exclude_k_var = tk.BooleanVar(value=self.cfg.get("exclude_prefix_k", True))
        ttk.Checkbutton(
            streit_frame,
            text="Exclude rows starting with 'K'",
            variable=self.exclude_k_var,
        ).pack(anchor="w", pady=(6, 0))

        self.ignore_top_doc_row_var = tk.BooleanVar(
            value=self.cfg.get("ignore_top_doc_row", False)
        )
        ttk.Checkbutton(
            streit_frame,
            text="Ignore first doc row match",
            variable=self.ignore_top_doc_row_var,
        ).pack(anchor="w")

        row3 = ttk.Frame(streit_frame)
        row3.pack(anchor="w", pady=(6, 0))
        ttk.Label(row3, text="PDF search term").pack(side=tk.LEFT)
        self.streitwort_var = tk.StringVar(
            value=self.cfg.get("streitwert_term", "Streitwert")
        )
        ttk.Entry(row3, textvariable=self.streitwort_var, width=16).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Label(row3, text="Open wait (s)").pack(side=tk.LEFT)
        self.docwait_var = tk.StringVar(value=str(self.cfg.get("doc_open_wait", 1.2)))
        ttk.Entry(row3, textvariable=self.docwait_var, width=6).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Label(row3, text="Hit wait (s)").pack(side=tk.LEFT)
        self.hitwait_var = tk.StringVar(value=str(self.cfg.get("pdf_hit_wait", 1.0)))
        ttk.Entry(row3, textvariable=self.hitwait_var, width=6).pack(
            side=tk.LEFT, padx=6
        )

        self.skip_waits_var = tk.BooleanVar(
            value=self.cfg.get("streitwert_overlay_skip_waits", False)
        )
        ttk.Checkbutton(
            streit_frame,
            text="Only wait for loading overlays (ignore manual delays)",
            variable=self.skip_waits_var,
        ).pack(anchor="w", pady=(6, 0))

        wait_row = ttk.Frame(streit_frame)
        wait_row.pack(anchor="w", pady=(2, 0))
        ttk.Label(wait_row, text="PDF view wait (s)").pack(side=tk.LEFT)
        self.pdf_view_wait_var = tk.StringVar(
            value=str(self.cfg.get("pdf_view_extra_wait", 2.0))
        )
        ttk.Entry(wait_row, textvariable=self.pdf_view_wait_var, width=6).pack(
            side=tk.LEFT, padx=6
        )

        ttk.Label(streit_frame, text="Streitwert CSV").pack(anchor="w", pady=(6, 0))
        self.streit_csv_var = tk.StringVar(
            value=self.cfg.get("streitwert_results_csv", "streitwert_results.csv")
        )
        ttk.Entry(streit_frame, textvariable=self.streit_csv_var, width=40).pack(
            anchor="w", pady=(0, 4)
        )

        ttk.Label(streit_frame, text="Streitwert_Results_Rechnungen.csv").pack(
            anchor="w", pady=(0, 0)
        )
        self.rechnungen_csv_var = tk.StringVar(
            value=self.cfg.get(
                "rechnungen_results_csv", "Streitwert_Results_Rechnungen.csv"
            )
        )
        ttk.Entry(streit_frame, textvariable=self.rechnungen_csv_var, width=40).pack(
            anchor="w", pady=(0, 4)
        )

        ttk.Button(
            streit_frame,
            text="Test Streitwert Setup",
            command=self.test_streitwert_setup,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Button(
            streit_frame,
            text="Test Instance Extraction",
            command=self.test_instance_extraction,
        ).pack(anchor="w", pady=(4, 0))

        ttk.Button(
            streit_frame,
            text="Run Streitwert Scan",
            command=self.run_streitwert_threaded,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Button(
            streit_frame,
            text="Start Streitwert Scan + Rechnungen",
            command=self.run_streitwert_with_rechnungen_threaded,
        ).pack(anchor="w", pady=(6, 0))

        # --- Rechnungen tab ---
        rechn_frame = ttk.LabelFrame(rechn_tab, text="Rechnungen Calibration & Test")
        rechn_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        ttk.Label(
            rechn_frame,
            text=(
                "Calibrate the Rechnungen list capture before combining with the "
                "Streitwert workflow."
            ),
        ).pack(anchor="w", pady=(0, 6))

        rcal = ttk.Frame(rechn_frame)
        rcal.pack(anchor="w", pady=(0, 2))
        ttk.Button(
            rcal,
            text="Pick Rechnungen Region",
            command=self.pick_rechnungen_region,
        ).pack(side=tk.LEFT, padx=2)

        rechn_box = self.cfg.get("rechnungen_region")
        if (
            isinstance(rechn_box, (list, tuple))
            and len(rechn_box) == 4
            and all(isinstance(v, (int, float)) for v in rechn_box)
        ):
            rechn_txt = (
                f"{rechn_box[0]:.3f}, {rechn_box[1]:.3f}, "
                f"{rechn_box[2]:.3f}, {rechn_box[3]:.3f}"
            )
        else:
            rechn_txt = ""
        self.rechnungen_region_var = tk.StringVar(value=rechn_txt)
        ttk.Entry(
            rechn_frame,
            textvariable=self.rechnungen_region_var,
            width=40,
            state="readonly",
        ).pack(anchor="w", pady=(0, 6))

        ttk.Button(
            rechn_frame,
            text="Pick GG Bezeichnung Region",
            command=self.pick_rechnungen_gg_region,
        ).pack(anchor="w", pady=(0, 2))

        gg_box = self.cfg.get("rechnungen_gg_region")
        if (
            isinstance(gg_box, (list, tuple))
            and len(gg_box) == 4
            and all(isinstance(v, (int, float)) for v in gg_box)
        ):
            gg_txt = (
                f"{gg_box[0]:.3f}, {gg_box[1]:.3f}, "
                f"{gg_box[2]:.3f}, {gg_box[3]:.3f}"
            )
        else:
            gg_txt = ""
        self.rechnungen_gg_region_var = tk.StringVar(value=gg_txt)
        ttk.Entry(
            rechn_frame,
            textvariable=self.rechnungen_gg_region_var,
            width=40,
            state="readonly",
        ).pack(anchor="w", pady=(0, 6))

        timing_box = ttk.LabelFrame(rechn_frame, text="Timing & Options")
        timing_box.pack(fill=tk.X, pady=(0, 6))

        search_row = ttk.Frame(timing_box)
        search_row.pack(anchor="w", pady=(0, 2))
        ttk.Label(search_row, text="Search wait (sec)").pack(side=tk.LEFT)
        self.rechnungen_search_wait_var = tk.StringVar(
            value=str(self.cfg.get("rechnungen_search_wait", 1.2))
        )
        ttk.Entry(
            search_row, textvariable=self.rechnungen_search_wait_var, width=6
        ).pack(side=tk.LEFT, padx=6)

        region_row = ttk.Frame(timing_box)
        region_row.pack(anchor="w", pady=(0, 2))
        ttk.Label(region_row, text="Region wait (sec)").pack(side=tk.LEFT)
        self.rechnungen_region_wait_var = tk.StringVar(
            value=str(self.cfg.get("rechnungen_region_wait", 0.8))
        )
        ttk.Entry(
            region_row, textvariable=self.rechnungen_region_wait_var, width=6
        ).pack(side=tk.LEFT, padx=6)

        self.rechnungen_skip_waits_var = tk.BooleanVar(
            value=self.cfg.get("rechnungen_overlay_skip_waits", False)
        )
        ttk.Checkbutton(
            timing_box,
            text="Only wait for loading overlays (ignore manual delays)",
            variable=self.rechnungen_skip_waits_var,
        ).pack(anchor="w", pady=(2, 0))

        ttk.Button(
            rechn_frame,
            text="Test Rechnungen Extraction",
            command=self.test_rechnungen_threaded,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Button(
            rechn_frame,
            text="Test GG Extraction",
            command=self.test_rechnungen_gg_threaded,
        ).pack(anchor="w", pady=(4, 0))

        ttk.Label(rechn_frame, text="Rechnungen-only CSV").pack(anchor="w", pady=(6, 0))
        self.rechnungen_only_csv_var = tk.StringVar(
            value=self.cfg.get(
                "rechnungen_only_results_csv", "rechnungen_only_results.csv"
            )
        )
        ttk.Entry(
            rechn_frame, textvariable=self.rechnungen_only_csv_var, width=40
        ).pack(anchor="w", pady=(0, 4))

        ttk.Button(
            rechn_frame,
            text="Run Rechnungen Extraction",
            command=self.run_rechnungen_only_threaded,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Label(rechn_frame, text="GG Extract CSV").pack(anchor="w", pady=(6, 0))
        self.rechnungen_gg_csv_var = tk.StringVar(
            value=self.cfg.get(
                "rechnungen_gg_results_csv", "rechnungen_gg_results.csv"
            )
        )
        ttk.Entry(
            rechn_frame, textvariable=self.rechnungen_gg_csv_var, width=40
        ).pack(anchor="w", pady=(0, 4))

        ttk.Button(
            rechn_frame,
            text="Run GG Extraction",
            command=self.run_rechnungen_gg_threaded,
        ).pack(anchor="w", pady=(6, 0))

        # --- Log tab ---
        log_frame = ttk.LabelFrame(log_tab, text="Streitwert Log Extraction")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        ttk.Label(log_frame, text="Log folder").pack(anchor="w")
        log_dir_row = ttk.Frame(log_frame)
        log_dir_row.pack(anchor="w", fill=tk.X, pady=(0, 4))
        self.log_dir_var = tk.StringVar(value=self.cfg.get("log_folder", LOG_DIR))
        ttk.Entry(
            log_dir_row,
            textvariable=self.log_dir_var,
            width=42,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(
            log_dir_row,
            text="Browse",
            command=self.browse_log_dir,
        ).pack(side=tk.LEFT, padx=6)

        ttk.Label(log_frame, text="Output CSV").pack(anchor="w")
        self.log_extract_csv_var = tk.StringVar(
            value=self.cfg.get("log_extract_results_csv", "streitwert_log_extract.csv")
        )
        ttk.Entry(log_frame, textvariable=self.log_extract_csv_var, width=40).pack(
            anchor="w", pady=(0, 4)
        )

        ttk.Button(
            log_frame,
            text="Extract Streitwert from Logs",
            command=self.run_log_extraction_threaded,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Label(
            log_frame,
            text=(
                "Parses saved OCR logs and records the first detected Streitwert "
                "amount per file."
            ),
            wraplength=320,
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(6, 0))

        # --- OCR / Profiles tab ---
        ocr_options = ttk.LabelFrame(ocr_tab, text="OCR Options")
        ocr_options.pack(fill=tk.X, pady=4)
        rowb = ttk.Frame(ocr_options)
        rowb.pack(anchor="w", pady=(0, 4))
        ttk.Label(rowb, text="Upscale ×").pack(side=tk.LEFT)
        self.upscale_var = tk.StringVar(value=str(self.cfg.get("upscale_x", 4)))
        ttk.Entry(rowb, textvariable=self.upscale_var, width=5).pack(
            side=tk.LEFT, padx=6
        )
        self.color_var = tk.BooleanVar(value=self.cfg.get("color_ocr", True))
        ttk.Checkbutton(rowb, text="Color OCR", variable=self.color_var).pack(
            side=tk.LEFT, padx=6
        )

        fr = ttk.Frame(ocr_options)
        fr.pack(anchor="w", pady=(4, 0))
        self.fullparse_var = tk.BooleanVar(
            value=self.cfg.get("use_full_region_parse", True)
        )
        ttk.Checkbutton(
            fr, text="Use full-region parsing", variable=self.fullparse_var
        ).pack(side=tk.LEFT)
        ttk.Label(fr, text="Keyword").pack(side=tk.LEFT, padx=(12, 4))
        self.keyword_var = tk.StringVar(value=self.cfg.get("keyword", "Honorar"))
        ttk.Entry(fr, textvariable=self.keyword_var, width=16).pack(side=tk.LEFT)

        nr = ttk.Frame(ocr_options)
        nr.pack(anchor="w", pady=(4, 0))
        self.normalize_var = tk.BooleanVar(value=self.cfg.get("normalize_ocr", True))
        ttk.Checkbutton(
            nr, text="Normalize OCR (O→0, S→5…)", variable=self.normalize_var
        ).pack(side=tk.LEFT)

        ttk.Button(
            ocr_options, text="Test Parse (full region)", command=self.test_parse_full
        ).pack(anchor="w", pady=(6, 0))

        profile_frame = ttk.LabelFrame(ocr_tab, text="Amount Region Profiles")
        profile_frame.pack(fill=tk.BOTH, expand=True, pady=6)
        prof_row1 = ttk.Frame(profile_frame)
        prof_row1.pack(anchor="w", pady=(0, 4))
        ttk.Label(prof_row1, text="Active").pack(side=tk.LEFT)
        self.profile_names = [p["name"] for p in self.cfg.get("amount_profiles", [])]
        self.profile_var = tk.StringVar(value=self.cfg.get("active_amount_profile", ""))
        self.profile_box = ttk.Combobox(
            prof_row1,
            textvariable=self.profile_var,
            values=self.profile_names,
            width=28,
            state="readonly",
        )
        self.profile_box.pack(side=tk.LEFT, padx=6)

        self.use_profile_var = tk.BooleanVar(
            value=self.cfg.get("use_amount_profile", False)
        )
        ttk.Checkbutton(
            prof_row1, text="Use profile region", variable=self.use_profile_var
        ).pack(side=tk.LEFT)

        prof_row2 = ttk.Frame(profile_frame)
        prof_row2.pack(anchor="w", pady=(0, 4))
        ttk.Label(prof_row2, text="Name").pack(side=tk.LEFT)
        self.new_prof_name_var = tk.StringVar(
            value=self.cfg.get("active_amount_profile", "")
        )
        ttk.Entry(prof_row2, textvariable=self.new_prof_name_var, width=20).pack(
            side=tk.LEFT, padx=6
        )

        ttk.Label(prof_row2, text="Keyword").pack(side=tk.LEFT)
        self.prof_keyword_var = tk.StringVar(value="")
        ttk.Entry(prof_row2, textvariable=self.prof_keyword_var, width=16).pack(
            side=tk.LEFT, padx=6
        )

        prof_row3 = ttk.Frame(profile_frame)
        prof_row3.pack(anchor="w", pady=(0, 4))
        ttk.Button(
            prof_row3, text="Pick Amount Region", command=self.pick_amount_region
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            prof_row3, text="New / Save Profile", command=self.save_amount_profile
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            prof_row3, text="Delete Profile", command=self.delete_amount_profile
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            prof_row3, text="Test Parse (profile)", command=self.test_parse_profile
        ).pack(side=tk.LEFT, padx=2)

        # --- Footer actions below notebook ---
        actions = ttk.Frame(left_container)
        actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(actions, text="Save Config", command=self.save_config).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(actions, text="Load Config", command=self.load_config).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(actions, text="Run Batch", command=self.run_batch_threaded).pack(
            side=tk.RIGHT, padx=2
        )

        # --- Right frame (preview + log) ---
        right = ttk.Frame(self, padding=10)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(right, text="OCR Preview (color crop)").pack(anchor="w")
        self.img_label = ttk.Label(right)
        self.img_label.pack(anchor="w", pady=(0, 6))
        ttk.Button(
            right, text="Toggle Live Preview", command=self.toggle_live_preview
        ).pack(anchor="w", pady=(0, 6))
        ttk.Label(right, text="Logs").pack(anchor="w")
        log_notebook = ttk.Notebook(right)
        log_notebook.pack(fill=tk.BOTH, expand=True)

        detailed_tab = ttk.Frame(log_notebook)
        simple_tab = ttk.Frame(log_notebook)
        log_notebook.add(detailed_tab, text="Detailed Log")
        log_notebook.add(simple_tab, text="Simple Log")

        detailed_frame = ttk.Frame(detailed_tab)
        detailed_frame.pack(fill=tk.BOTH, expand=True)
        self.log = tk.Text(detailed_frame, height=14, wrap="word")
        self.log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll = ttk.Scrollbar(
            detailed_frame, orient="vertical", command=self.log.yview
        )
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.configure(yscrollcommand=log_scroll.set)

        simple_frame = ttk.Frame(simple_tab)
        simple_frame.pack(fill=tk.BOTH, expand=True)
        self.simple_log = tk.Text(
            simple_frame, height=14, wrap="word", state="disabled"
        )
        self.simple_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        simple_scroll = ttk.Scrollbar(
            simple_frame, orient="vertical", command=self.simple_log.yview
        )
        simple_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.simple_log.configure(yscrollcommand=simple_scroll.set)

        # Load active profile details into fields
        self._refresh_profile_fields_from_active()

    # ---------- UI helpers ----------
    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.xls_var.set(path)

    def browse_tesseract(self):
        path = filedialog.askopenfilename(
            title="Locate tesseract.exe (or choose its folder)",
            filetypes=[("tesseract.exe", "tesseract.exe"), ("All files", "*.*")],
        )
        if path:
            self.tess_var.set(path)

    def browse_log_dir(self):
        path = filedialog.askdirectory(title="Select log folder")
        if path:
            self.log_dir_var.set(path)

    def connect_rdp(self):
        try:
            win, rect = connect_rdp_window(self.rdp_var.get())
            self.current_rect = rect
            self.log_print(f"Connected. Client rect: {rect}")
        except Exception as e:
            messagebox.showerror("Connect RDP", f"Failed: {e}")

    def _show_capture_countdown(self, seconds=None, message="Capturing in {n}…"):
        secs = int(seconds if seconds is not None else self.capture_countdown_seconds)
        if secs <= 0:
            return
        try:
            top = tk.Toplevel(self)
        except tk.TclError:
            time.sleep(secs)
            return

        top.title("Countdown")
        top.transient(self)
        top.attributes("-topmost", True)
        try:
            x = self.winfo_rootx() + 120
            y = self.winfo_rooty() + 120
            top.geometry(f"200x90+{x}+{y}")
        except Exception:
            top.geometry("200x90")

        label = ttk.Label(top, text="", padding=16, anchor="center")
        label.pack(fill=tk.BOTH, expand=True)
        top.update()

        try:
            for remaining in range(secs, 0, -1):
                label.configure(text=message.format(n=remaining))
                top.update()
                time.sleep(1)
        finally:
            try:
                top.destroy()
            except tk.TclError:
                pass

    def _prompt_and_capture_point(self, title, prompt, seconds=None):
        msg = f"{prompt}\n\nCountdown starts after you click OK."
        messagebox.showinfo(title, msg)
        self._show_capture_countdown(seconds=seconds)
        return pyautogui.position()

    def pick_search_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick Search Point", "Position your mouse over the search bar location."
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["search_point"] = rel
        self.sp_var.set(f"{rel[0]:.3f}, {rel[1]:.3f}")
        self.log_print(f"Search point set (relative): {rel}")

    def pick_result_region(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x1, y1 = self._prompt_and_capture_point(
            "Pick Result Region",
            "Position your mouse over the TOP-LEFT corner of the region.",
        )

        x2, y2 = self._prompt_and_capture_point(
            "Pick Result Region",
            "Now position your mouse over the BOTTOM-RIGHT corner.",
        )
        left, top = min(x1, x2), min(y1, y2)
        width, height = abs(x2 - x1), abs(y2 - y1)
        rel_box = abs_to_rel(self.current_rect, abs_box=(left, top, width, height))
        self.cfg["result_region"] = rel_box
        self.rr_var.set(
            f"{rel_box[0]:.3f}, {rel_box[1]:.3f}, {rel_box[2]:.3f}, {rel_box[3]:.3f}"
        )
        self.log_print(f"Result region set (relative): {rel_box}")

    def _two_click_box(self, msg1, msg2):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return None
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x1, y1 = self._prompt_and_capture_point("Pick", msg1)
        x2, y2 = self._prompt_and_capture_point("Pick", msg2)
        left, top = min(x1, x2), min(y1, y2)
        width, height = abs(x2 - x1), abs(y2 - y1)
        return abs_to_rel(self.current_rect, abs_box=(left, top, width, height))

    def pick_doclist_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the document list area, then OK.",
            "Hover BOTTOM-RIGHT of the document list area, then OK.",
        )
        if rb:
            self.cfg["doclist_region"] = rb
            self.log_print(f"Doc list region set: {rb}")

    def pick_rechnungen_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the Rechnungen list, then OK.",
            "Hover BOTTOM-RIGHT of the Rechnungen list, then OK.",
        )
        if rb:
            self.cfg["rechnungen_region"] = rb
            if hasattr(self, "rechnungen_region_var"):
                self.rechnungen_region_var.set(
                    f"{rb[0]:.3f}, {rb[1]:.3f}, {rb[2]:.3f}, {rb[3]:.3f}"
                )
            self.log_print(f"Rechnungen region set: {rb}")

    def pick_rechnungen_gg_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the GG Bezeichnung area, then OK.",
            "Hover BOTTOM-RIGHT of the GG area, then OK.",
        )
        if rb:
            self.cfg["rechnungen_gg_region"] = rb
            if hasattr(self, "rechnungen_gg_region_var"):
                self.rechnungen_gg_region_var.set(
                    f"{rb[0]:.3f}, {rb[1]:.3f}, {rb[2]:.3f}, {rb[3]:.3f}"
                )
            self.log_print(f"GG region set: {rb}")

    def pick_fees_file_search_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the FILE-SEARCH input, then OK.",
            "Hover BOTTOM-RIGHT of the FILE-SEARCH input, then OK.",
        )
        if rb:
            self.cfg["fees_file_search_region"] = rb
            self.log_print(f"Fees file-search region set: {rb}")

    def pick_fees_seiten_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the Seiten (thumbnails) strip, then OK.",
            "Hover BOTTOM-RIGHT of the Seiten strip, then OK.",
        )
        if rb:
            self.cfg["fees_seiten_region"] = rb
            self.log_print(f"Fees Seiten region set: {rb}")

    def pick_instance_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the AZ Instanz table (include the header row). Then OK.",
            "Hover BOTTOM-RIGHT (also include a bit of space below headers). Then OK.",
        )
        if rb:
            self.cfg["instance_region"] = rb
            if hasattr(self, "instance_var"):
                self.instance_var.set(
                    f"{rb[0]:.3f}, {rb[1]:.3f}, {rb[2]:.3f}, {rb[3]:.3f}"
                )
            self.log_print(f"Instance table region set: {rb}")

    def pick_pdf_hits_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick",
            "Hover the PDF results button you want to click (top match), then confirm.",
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["pdf_hits_point"] = rel
        if hasattr(self, "pdf_hits_var"):
            self.pdf_hits_var.set(f"{rel[0]:.3f}, {rel[1]:.3f}")
        self.log_print(f"PDF hits button point set: {rel}")

    def pick_pdf_second_hits_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick",
            "Hover the PDF results button for the SECOND match, then confirm.",
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["pdf_hits_second_point"] = rel
        if hasattr(self, "pdf_hits2_var"):
            self.pdf_hits2_var.set(f"{rel[0]:.3f}, {rel[1]:.3f}")
        self.log_print(f"Secondary PDF hits button point set: {rel}")

    def pick_pdf_third_hits_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick",
            "Hover the PDF results button for the THIRD match, then confirm.",
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["pdf_hits_third_point"] = rel
        if hasattr(self, "pdf_hits3_var"):
            self.pdf_hits3_var.set(f"{rel[0]:.3f}, {rel[1]:.3f}")
        self.log_print(f"Tertiary PDF hits button point set: {rel}")

    def pick_pdf_text_region(self):
        rb = self._two_click_box(
            "Hover TOP-LEFT of the PDF page text area, then OK.",
            "Hover BOTTOM-RIGHT of the PDF page text area, then OK.",
        )
        if rb:
            self.cfg["pdf_text_region"] = rb
            self.log_print(f"PDF text region set: {rb}")

    def pick_pdf_search_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick", "Hover the PDF search box caret position, then confirm."
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["pdf_search_point"] = rel
        self.log_print(f"PDF search point set: {rel}")

    def pick_doc_view_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick",
            "Hover the View button used to open documents, then confirm.",
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["doc_view_point"] = rel
        self.doc_view_var.set(f"{rel[0]:.3f}, {rel[1]:.3f}")
        self.log_print(f"View button point set: {rel}")

    def pick_pdf_close_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x, y = self._prompt_and_capture_point(
            "Pick",
            "Hover the PDF window close button (X), then confirm.",
        )
        rel = abs_to_rel(self.current_rect, abs_point=(x, y))
        self.cfg["pdf_close_point"] = rel
        if hasattr(self, "pdf_close_var"):
            self.pdf_close_var.set(f"{rel[0]:.3f}, {rel[1]:.3f}")
        self.log_print(f"PDF close button point set: {rel}")

    def pick_amount_region(self):
        """Pick a sub-region INSIDE the current Result Region; saves it into the profile editor fields."""
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        if not self.cfg.get("result_region"):
            messagebox.showwarning(
                "Pick Amount Region", "Please set the Result Region first."
            )
            return

        # Absolute rectangle of the current result_region
        rx, ry, rw, rh = rel_to_abs(self.current_rect, self.cfg["result_region"])
        outer_abs = (rx, ry, rx + rw, ry + rh)

        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        x1, y1 = self._prompt_and_capture_point(
            "Pick Amount Region",
            "Place your mouse at the TOP-LEFT of the amount area inside the Result Region.",
        )

        x2, y2 = self._prompt_and_capture_point(
            "Pick Amount Region", "Now move to the BOTTOM-RIGHT corner."
        )

        # Clamp to outer region and convert to sub-relative
        left, top = max(min(x1, x2), outer_abs[0]), max(min(y1, y2), outer_abs[1])
        right, bottom = min(max(x1, x2), outer_abs[2]), min(max(y1, y2), outer_abs[3])
        width, height = max(1, right - left), max(1, bottom - top)

        # Compute relative to outer (result_region) box
        sub_rel = [
            (left - outer_abs[0]) / (outer_abs[2] - outer_abs[0]),
            (top - outer_abs[1]) / (outer_abs[3] - outer_abs[1]),
            width / (outer_abs[2] - outer_abs[0]),
            height / (outer_abs[3] - outer_abs[1]),
        ]
        self._current_profile_sub_region = sub_rel
        self.log_print(
            f"Picked Amount Sub-Region (relative to result_region): {', '.join(f'{v:.3f}' for v in sub_rel)}"
        )

    def test_typing(self):
        try:
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
            x, y = rel_to_abs(self.current_rect, self.cfg["search_point"])
            pyautogui.click(x, y)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            txt = self.typing_test_var.get() or "TEST123"
            pyautogui.typewrite(txt, interval=float(self.type_var.get() or 0.02))
            pyautogui.press("enter")
            self.log_print(f"Typed test text: {txt}")
        except Exception as e:
            messagebox.showerror("Test Typing", f"Failed: {e}")

    def test_parse_full(self):
        try:
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            text, crop_img, lines, amount = self.read_region_and_parse(
                use_profile=False
            )
            self.show_preview(crop_img)
            self.log_print(
                "OCR preview (full region text):\n" + (text if text else "(no text)")
            )
            self.log_print(f"Extracted amount: {amount or '(none)'}")
        except Exception as e:
            messagebox.showerror("Test Parse", f"Failed: {e}")

    def test_parse_profile(self):
        try:
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            text, crop_img, lines, amount = self.read_region_and_parse(use_profile=True)
            self.show_preview(crop_img)
            self.log_print(
                "OCR preview (profile sub-region text):\n"
                + (text if text else "(no text)")
            )
            self.log_print(f"Extracted amount: {amount or '(none)'}")
        except Exception as e:
            messagebox.showerror("Test Parse (profile)", f"Failed: {e}")

    # ---------- Core OCR/parse ----------
    def apply_paths_to_tesseract(self):
        tp = (self.tess_var.get() or "").strip().strip('"').replace("/", "\\")
        if not tp.lower().endswith("tesseract.exe"):
            if os.path.isdir(tp):
                tp = os.path.join(tp, "tesseract.exe")
            else:
                tp = tp + ("" if tp.lower().endswith(".exe") else "\\tesseract.exe")
        if not os.path.exists(tp):
            raise FileNotFoundError(f"Tesseract not found at: {tp}")
        pytesseract.pytesseract.tesseract_cmd = tp
        self.log_print(f"Tesseract path set: {tp}")

    def _validate_result_region(self):
        l, t, w, h = self.cfg["result_region"]
        if w <= 0 or h <= 0:
            raise ValueError(
                f"Invalid Result Region (w={w:.3f}, h={h:.3f}). "
                "Please click 'Pick Result Region' to set a non-zero area."
            )

    def _grab_region_color(self, x=None, y=None, w=None, h=None, upscale_x=None):
        """
        Grab a region of the screen in color and return it as a PIL Image.

        If x,y,w,h are provided, use those coordinates directly.
        Otherwise, use the configured result region.

        upscale_x: Optional scale factor. If not provided, uses self.upscale_var.
        """
        if x is None:  # Use result region
            self._validate_result_region()
            x, y, w, h = rel_to_abs(self.current_rect, self.cfg["result_region"])

        # Convert coordinates to integers for mss
        x, y = int(x), int(y)
        w, h = int(w), int(h)

        region = grab_xywh(x, y, w, h)
        scale = max(
            1, int(upscale_x if upscale_x is not None else self.upscale_var.get() or 3)
        )
        return upscale_pil(region, scale=scale)

    def _safe_crop(self, img, box):
        """
        Crop with clamping; guarantees at least 1x1 output.
        """
        x0, y0, x1, y1 = [int(v) for v in box]
        # Clamp to image bounds and enforce at least 1px size
        x0 = max(0, min(x0, img.width - 1))
        y0 = max(0, min(y0, img.height - 1))
        x1 = max(x0 + 1, min(int(x1), img.width))
        y1 = max(y0 + 1, min(int(y1), img.height))
        return img.crop((x0, y0, x1, y1))

    def _safe_save_png(self, img, path):
        """
        Save preview defensively. Skips saving if size is invalid.
        """
        try:
            w, h = img.size
            if w < 1 or h < 1:
                self.log_print(f"[Preview] skip save: empty image ({w}x{h}) -> {path}")
                return
            if img.mode not in ("RGB", "RGBA"):
                img = img.convert("RGB")
            tmp = path + ".tmp"
            img.save(tmp, format="PNG")
            os.replace(tmp, path)
        except Exception as e:
            try:
                self.log_print(f"[Preview] save failed: {e}")
            except Exception:
                pass

    def _crop_to_profile(self, img):
        """Crop the passed PIL image (which is the full Result Region) to the active profile's sub-region."""
        prof = self._get_active_profile()
        if not prof or not prof.get("sub_region"):
            return img
        l, t, w, h = prof["sub_region"]
        px_l = int(l * img.width)
        px_t = int(t * img.height)
        px_r = int((l + w) * img.width)
        px_b = int((t + h) * img.height)
        px_l = max(0, min(px_l, img.width - 1))
        px_t = max(0, min(px_t, img.height - 1))
        px_r = max(px_l + 1, min(px_r, img.width))
        px_b = max(px_t + 1, min(px_b, img.height))
        return img.crop((px_l, px_t, px_r, px_b))

    def read_region_and_parse(self, use_profile=False):
        """OCR selected region (full or profile sub-region), parse lines, then extract amount with keyword preference."""
        crop = self._grab_region_color()
        keyword = self.keyword_var.get().strip()
        if use_profile and self.use_profile_var.get():
            prof = self._get_active_profile()
            if prof:
                crop = self._crop_to_profile(crop)
                keyword = prof.get("keyword", keyword) or keyword

        lang = self.lang_var.get().strip() or "deu+eng"
        df = do_ocr_data(crop, lang=lang, psm=6)
        lines = lines_from_tsv(df)
        simple_lines = [(y, text) for _, y, _, _, text in lines]
        full_text = "\n".join(t for _, t in simple_lines)
        if self.normalize_var.get():
            normalized_lines = [(y, normalize_line_soft(t)) for y, t in simple_lines]
            full_text = "\n".join(t for _, t in normalized_lines)
        else:
            normalized_lines = simple_lines
        amount, line = extract_amount_from_lines(normalized_lines, keyword=keyword)
        return full_text, crop, normalized_lines, amount

    # ---------- Batch ----------
    def run_batch_threaded(self):
        threading.Thread(target=self.run_batch, daemon=True).start()

    def run_batch(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()

            _, rect = connect_rdp_window(self.rdp_var.get())
            self.current_rect = rect

            df = pd.read_excel(
                self.cfg["excel_path"], sheet_name=self.cfg["excel_sheet"]
            )
            results = []
            start_cell = (self.start_cell_var.get() or "").strip()
            max_rows = int(self.max_rows_var.get() or "0")

            # Determine column & start row
            if start_cell:
                m = re.match(r"^\s*([A-Za-z]+)\s*([0-9]+)\s*$", start_cell)
                if not m:
                    self.log_print(
                        f"ERROR: Invalid start cell '{start_cell}'. Use like 'B2'."
                    )
                    return
                col_letters, row_num = m.group(1).upper(), int(m.group(2))
                col_idx = 0
                for ch in col_letters:
                    col_idx = col_idx * 26 + (ord(ch) - 64)
                col_idx -= 1
                i0 = max(row_num - 2, 0)
                rows = df.iloc[i0:]
            else:
                if self.cfg["input_column"] not in df.columns:
                    self.log_print(
                        f"ERROR: column '{self.cfg['input_column']}' not in sheet."
                    )
                    return
                col_idx = df.columns.get_loc(self.cfg["input_column"])
                rows = df

            if max_rows > 0:
                rows = rows.head(max_rows)
            total = len(rows)

            for _, row in rows.iterrows():
                q = str(row.iloc[col_idx])

                # Type query into search box
                Desktop(backend="uia").window(
                    title_re=self.cfg["rdp_title_regex"]
                ).set_focus()
                x, y = rel_to_abs(self.current_rect, self.cfg["search_point"])
                pyautogui.click(x, y)
                pyautogui.hotkey("ctrl", "a")
                pyautogui.press("backspace")
                pyautogui.typewrite(q, interval=float(self.cfg["type_delay"]))
                pyautogui.press("enter")
                time.sleep(float(self.cfg["post_search_wait"]))

                # Parse (profile region if enabled, else full region)
                use_prof = bool(self.cfg.get("use_amount_profile", False))
                full_text, _crop, lines, amount = self.read_region_and_parse(
                    use_profile=use_prof
                )

                rec = row.to_dict()
                rec["__query__"] = q
                rec["extracted_text"] = full_text
                rec["extracted_amount"] = amount
                rec["extracted_line"] = next(
                    (t for _, t in lines if amount and amount in t), ""
                )
                results.append(rec)
                self.log_print(f"[{len(results)}/{total}] {q} → {amount or '(none)'}")

            out = pd.DataFrame(results)
            out.to_csv(self.cfg["results_csv"], index=False, encoding="utf-8-sig")
            self.log_print(f"Done. Saved to {self.cfg['results_csv']}")

        except Exception as e:
            self.log_print("ERROR: " + repr(e))

    def test_rechnungen_threaded(self):
        t = threading.Thread(target=self.test_rechnungen, daemon=True)
        t.start()

    def test_rechnungen(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()

            inst_info = self.detect_instance(prefix="[Rechnungen Test] ")
            summary = self._extract_rechnungen_summary(prefix="[Rechnungen Test] ")
            if summary is None:
                self.log_print("[Rechnungen Test] No Rechnungen data detected.")
                return
            self._log_rechnungen_summary("[Rechnungen Test] ", summary)
        except Exception as e:
            self.log_print(f"[Rechnungen Test] ERROR: {e!r}")

    def test_rechnungen_gg_threaded(self):
        t = threading.Thread(target=self.test_rechnungen_gg, daemon=True)
        t.start()

    def test_rechnungen_gg(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()

            self.clear_simple_log()
            entries = self._extract_rechnungen_gg_entries(prefix="[GG Test] ")
            self._wait_for_preview_ready()
            if not entries:
                self.log_print("[GG Test] No GG transactions detected.")
                self.simple_log_print("GG Test: no GG transactions detected.")
                return
            summary_parts = []
            for idx, entry in enumerate(entries, 1):
                detail = self._format_rechnungen_detail(entry)
                amount = entry.get("amount", "") or "(no amount)"
                self.log_print(f"[GG Test] #{idx}: {amount}{detail}")
                summary_parts.append(f"{amount}{detail}")
            if summary_parts:
                self.simple_log_print(f"GG Test: {'; '.join(summary_parts)}")
        except Exception as e:
            self.log_print(f"[GG Test] ERROR: {e!r}")

    def run_rechnungen_only_threaded(self):
        t = threading.Thread(target=self.run_rechnungen_only, daemon=True)
        t.start()

    def run_rechnungen_only(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()

            if not self._doclist_abs_rect():
                self.log_print(
                    "Doc list region is not configured. Please re-run calibration."
                )
                return

            queries = self._gather_aktenzeichen()
            if not queries:
                return

            if queries:
                self.clear_simple_log()
            skip_waits = self._should_skip_manual_waits()
            wait_setting = self.cfg.get(
                "rechnungen_search_wait", self.cfg.get("post_search_wait", 1.2)
            )
            try:
                wait_seconds = float(wait_setting)
            except Exception:
                wait_seconds = float(DEFAULTS.get("rechnungen_search_wait", 1.2))
            list_wait = 0.0 if skip_waits else max(0.0, wait_seconds)

            results = []
            total = len(queries)
            for idx, (aktenzeichen, _row) in enumerate(queries, 1):
                prefix = f"[Rechnungen {idx}/{total}] "
                self.log_print(
                    f"{prefix}Searching doc list for Aktenzeichen: {aktenzeichen}"
                )
                inst_info = None
                if not self._type_doclist_query(aktenzeichen, prefix=prefix):
                    self.log_print(
                        f"{prefix}Unable to type Aktenzeichen. Skipping entry."
                    )
                    continue
                if list_wait > 0:
                    time.sleep(list_wait)
                self.log_print(
                    f"{prefix}Typed '{aktenzeichen}' into the document search box."
                )
                inst_info = self.detect_instance(prefix=prefix)

                self._wait_for_doc_search_ready(
                    prefix=prefix, reason="after Aktenzeichen search"
                )
                self._wait_for_doclist_ready(
                    prefix=prefix, reason="after Aktenzeichen search"
                )

                summary = self._extract_rechnungen_summary(prefix=prefix)
                if summary is None:
                    self.log_print(
                        f"{prefix}Rechnungen capture returned no data; storing defaults."
                    )
                    summary = self._summarize_rechnungen_entries([])
                else:
                    self._log_rechnungen_summary(prefix, summary)

                row = self._build_rechnungen_result_row(aktenzeichen, summary)
                row["instance_detected"] = (inst_info or {}).get("instance")
                results.append(row)

            if results:
                pd.DataFrame(results).to_csv(
                    self.rechnungen_only_csv_var.get(),
                    index=False,
                    encoding="utf-8-sig",
                )
                self.log_print(
                    "Done. Saved Rechnungen-only results to "
                    f"{self.rechnungen_only_csv_var.get()}"
                )
            else:
                self.log_print(
                    "No Rechnungen values were captured from the Excel list."
                )

        except Exception as e:
            self.log_print(f"[Rechnungen] ERROR: {e!r}")

    def run_rechnungen_gg_threaded(self):
        t = threading.Thread(target=self.run_rechnungen_gg, daemon=True)
        t.start()

    def run_rechnungen_gg(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()

            if not self._doclist_abs_rect():
                self.log_print("Doc list region is not configured. Please re-run calibration.")
                return

            queries = self._gather_aktenzeichen()
            if not queries:
                return

            if queries:
                self.clear_simple_log()
            simple_lines = []

            skip_waits = self._should_skip_manual_waits()
            wait_setting = self.cfg.get(
                "rechnungen_search_wait", self.cfg.get("post_search_wait", 1.2)
            )
            try:
                wait_seconds = float(wait_setting)
            except Exception:
                wait_seconds = float(DEFAULTS.get("rechnungen_search_wait", 1.2))
            list_wait = 0.0 if skip_waits else max(0.0, wait_seconds)

            results = []
            total = len(queries)
            for idx, (aktenzeichen, _row) in enumerate(queries, 1):
                prefix = f"[GG {idx}/{total}] "
                simple_lines.append(f"{aktenzeichen}: (capturing GG...)")
                self._render_simple_log_lines(simple_lines)
                try:
                    self.log_print(
                        f"{prefix}Searching doc list for Aktenzeichen: {aktenzeichen}"
                    )
                    if not self._type_doclist_query(aktenzeichen, prefix=prefix):
                        self.log_print(
                            f"{prefix}Unable to type Aktenzeichen. Skipping entry."
                        )
                        simple_lines[-1] = f"{aktenzeichen}: (search failed)"
                        self._render_simple_log_lines(simple_lines)
                        continue
                    if list_wait > 0:
                        time.sleep(list_wait)
                    self.log_print(
                        f"{prefix}Typed '{aktenzeichen}' into the document search box."
                    )
                    self._wait_for_doc_search_ready(
                        prefix=prefix, reason="after Aktenzeichen search"
                    )
                    self._wait_for_doclist_ready(
                        prefix=prefix, reason="after Aktenzeichen search"
                    )

                    entries = self._extract_rechnungen_gg_entries(prefix=prefix)
                    if not self._wait_for_preview_ready(timeout=6.0):
                        self.log_print(
                            f"{prefix}Preview generation timed out; continuing with latest data."
                        )
                    if not entries:
                        self.log_print(f"{prefix}No GG transactions detected.")
                        summary_line = self._build_gg_summary_line(aktenzeichen, [])
                        simple_lines[-1] = summary_line
                        self._render_simple_log_lines(simple_lines)
                        self.log_print(f"{prefix}{summary_line}")
                        results.append(
                            {
                                "aktenzeichen": aktenzeichen,
                                "gg_detected": False,
                                "gg_count": 0,
                                "gg_amounts": "",
                                "gg_dates": "",
                                "gg_invoices": "",
                                "gg_raw": "",
                            }
                        )
                        continue

                    amounts = [entry.get("amount", "") or "" for entry in entries]
                    dates = [entry.get("date", "") or "" for entry in entries]
                    invoices = [entry.get("invoice", "") or "" for entry in entries]
                    raw_rows = [entry.get("raw", "") or "" for entry in entries]

                    for entry_idx, entry in enumerate(entries, 1):
                        detail = self._format_rechnungen_detail(entry)
                        amount = entry.get("amount", "") or "(no amount)"
                        self.log_print(f"{prefix}#{entry_idx}: {amount}{detail}")

                    results.append(
                        {
                            "aktenzeichen": aktenzeichen,
                            "gg_detected": True,
                            "gg_count": len(entries),
                            "gg_amounts": "; ".join(filter(None, amounts)),
                            "gg_dates": "; ".join(filter(None, dates)),
                            "gg_invoices": "; ".join(filter(None, invoices)),
                            "gg_raw": " || ".join(filter(None, raw_rows)),
                        }
                    )

                    summary_line = self._build_gg_summary_line(aktenzeichen, entries)
                    simple_lines[-1] = summary_line
                    self._render_simple_log_lines(simple_lines)
                    self.log_print(f"{prefix}{summary_line}")
                except Exception:
                    simple_lines[-1] = f"{aktenzeichen}: (error)"
                    self._render_simple_log_lines(simple_lines)
                    raise

            if results:
                pd.DataFrame(results).to_csv(
                    self.rechnungen_gg_csv_var.get(),
                    index=False,
                    encoding="utf-8-sig",
                )
                self.log_print(
                    "Done. Saved GG extraction results to "
                    f"{self.rechnungen_gg_csv_var.get()}"
                )
            else:
                self.log_print("No GG transactions were captured from the Excel list.")

        except Exception as e:
            self.log_print(f"[GG Extraction] ERROR: {e!r}")

    def run_log_extraction_threaded(self):
        t = threading.Thread(target=self.run_log_extraction, daemon=True)
        t.start()

    def run_log_extraction(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)

            log_dir = (self.log_dir_var.get() or "").strip() or LOG_DIR
            if not os.path.isdir(log_dir):
                ensure_log_dir()
                if not os.path.isdir(log_dir):
                    self.log_print(f"[Log Extract] Log directory not found: {log_dir}")
                    return

            files = [
                os.path.join(log_dir, name)
                for name in sorted(os.listdir(log_dir))
                if name.lower().endswith(".log")
            ]
            if not files:
                self.log_print(
                    f"[Log Extract] No .log files found in {os.path.abspath(log_dir)}"
                )
                return

            output_csv = (self.log_extract_csv_var.get() or "").strip()
            self.log_print(
                f"[Log Extract] Processing {len(files)} log file(s) from {os.path.abspath(log_dir)}"
            )

            results = []
            fallback_keywords = build_streitwert_keywords(
                self.streitwort_var.get().strip() or "Streitwert"
            )

            for path in files:
                label = self._log_label_from_filename(os.path.basename(path))
                amount, context, section = self._extract_amount_from_log(
                    path, fallback_keywords
                )
                display_label = label or os.path.basename(path)
                if amount:
                    results.append(
                        {
                            "log_file": os.path.basename(path),
                            "label": label,
                            "amount": amount,
                            "section": section or "",
                            "context": context or "",
                        }
                    )
                    self.log_print(
                        f"[Log Extract] {display_label} → {amount} ({section or 'section unknown'})"
                    )
                else:
                    self.log_print(f"[Log Extract] {display_label} → (none)")

            if results and output_csv:
                try:
                    pd.DataFrame(results).to_csv(
                        output_csv, index=False, encoding="utf-8-sig"
                    )
                    self.log_print(
                        f"[Log Extract] Saved {len(results)} entries to {output_csv}"
                    )
                except Exception as exc:
                    self.log_print(
                        f"[Log Extract] Failed to write CSV '{output_csv}': {exc}"
                    )
            elif results:
                self.log_print(
                    f"[Log Extract] Collected {len(results)} entries (CSV output disabled)."
                )
            else:
                self.log_print(
                    "[Log Extract] No Streitwert amounts were detected in the logs."
                )
        except Exception as e:
            self.log_print(f"[Log Extract] ERROR: {e}")

    def _log_label_from_filename(self, filename):
        base = os.path.splitext(filename)[0]
        m = re.match(r"^\d{8}-\d{6}_(.+)$", base)
        if m:
            base = m.group(1)
        label = base.replace("__", " – ")
        label = label.replace("_", " ")
        return label.strip()

    def _parse_log_file_sections(self, path):
        sections = []
        current_section = ""
        current_entries = []
        current_keywords = []
        last_coords = (0, 0, 0, 0)

        try:
            with open(path, "r", encoding="utf-8") as fh:
                for raw_line in fh:
                    line = raw_line.rstrip("\n")
                    stripped = line.strip()
                    if not stripped:
                        continue

                    sec_match = LOG_SECTION_RE.match(stripped)
                    if sec_match:
                        if current_entries:
                            sections.append(
                                {
                                    "section": current_section,
                                    "entries": list(current_entries),
                                    "keywords": list(current_keywords),
                                }
                            )
                        current_section = sec_match.group(1).strip()
                        current_entries = []
                        current_keywords = []
                        last_coords = (0, 0, 0, 0)
                        continue

                    kw_match = LOG_KEYWORD_RE.match(stripped)
                    if kw_match:
                        keywords = [
                            k.strip() for k in kw_match.group(1).split(",") if k.strip()
                        ]
                        current_keywords = keywords
                        continue

                    entry_match = LOG_ENTRY_RE.match(line)
                    if entry_match:
                        coords_text = entry_match.group(2)
                        coords_parts = [c.strip() for c in coords_text.split(",")]
                        if len(coords_parts) >= 4:
                            try:
                                x = int(coords_parts[0])
                                y = int(coords_parts[1])
                                w = int(coords_parts[2])
                                h = int(coords_parts[3])
                            except Exception:
                                x = y = w = h = 0
                        else:
                            x = y = w = h = 0
                        text = entry_match.group(3).strip()
                        current_entries.append((x, y, w, h, text))
                        last_coords = (x, y, w, h)
                        continue

                    soft_match = LOG_SOFT_RE.match(line)
                    if soft_match and current_entries:
                        text = soft_match.group(1).strip()
                        if text:
                            x, y, w, h = last_coords
                            current_entries.append((x, y, w, h, text))
                        continue

                    norm_match = LOG_NORM_RE.match(line)
                    if norm_match and current_entries:
                        text = norm_match.group(1).strip()
                        if text:
                            x, y, w, h = last_coords
                            current_entries.append((x, y, w, h, text))
                        continue
        except Exception as exc:
            self.log_print(f"[Log Extract] Failed to read log '{path}': {exc}")
            return []

        if current_entries:
            sections.append(
                {
                    "section": current_section,
                    "entries": list(current_entries),
                    "keywords": list(current_keywords),
                }
            )

        return sections

    def _extract_amount_from_log(self, path, fallback_keywords):
        sections = self._parse_log_file_sections(path)
        if not sections:
            return None, None, None

        for info in sections:
            entries = info.get("entries") or []
            if not entries:
                continue
            keywords = info.get("keywords") or fallback_keywords
            amount, line = extract_amount_from_lines(
                entries,
                keyword=keywords,
                min_value=STREITWERT_MIN_AMOUNT,
            )
            if not amount and keywords is not fallback_keywords:
                amount, line = extract_amount_from_lines(
                    entries,
                    keyword=fallback_keywords,
                    min_value=STREITWERT_MIN_AMOUNT,
                )
            if amount:
                return clean_amount_display(amount), line, info.get("section")

        combined_entries = []
        for info in sections:
            combined_entries.extend(info.get("entries") or [])

        if combined_entries:
            amount, line = extract_amount_from_lines(
                combined_entries,
                keyword=fallback_keywords,
                min_value=STREITWERT_MIN_AMOUNT,
            )
            if amount:
                return clean_amount_display(amount), line, "combined"

        return None, None, None

    def run_streitwert_threaded(self):
        t = threading.Thread(target=self.run_streitwert, daemon=True)
        t.start()

    def run_streitwert_with_rechnungen_threaded(self):
        t = threading.Thread(target=self.run_streitwert_with_rechnungen, daemon=True)
        t.start()

    def run_streitwert_with_rechnungen(self):
        self.run_streitwert(include_rechnungen=True)

    def _filter_streitwert_rows(self, lines):
        inc = [
            t.strip().lower()
            for t in (self.includes_var.get() or "").split(",")
            if t.strip()
        ]
        inc_match = [(tok, normalize_for_token_match(tok)) for tok in inc]
        exc = [
            t.strip().lower()
            for t in (self.excludes_var.get() or "").split(",")
            if t.strip()
        ]
        exc_match = [(tok, normalize_for_token_match(tok)) for tok in exc]
        excl_k = bool(self.exclude_k_var.get())

        matches = []
        debug_rows = []
        for x, y, w, h, txt in lines:
            raw = (txt or "").strip()
            if not raw:
                continue
            norm = normalize_line(raw)
            low_raw = raw.lower()
            low_norm = norm.lower()
            soft_raw = normalize_for_token_match(raw)
            soft_norm = normalize_for_token_match(norm)
            forced_skipped = False
            for label, pattern in FORCED_STREITWERT_EXCLUDES:
                try:
                    if pattern.search(raw):
                        debug_rows.append((raw, f"forced exclude '{label}'"))
                        forced_skipped = True
                        break
                except Exception:
                    continue
            if forced_skipped:
                continue
            if excl_k and re.match(r"^\s*k", low_raw):
                debug_rows.append((raw, "excluded prefix 'K'"))
                continue
            if exc:
                fields = [f for f in (low_raw, low_norm, soft_raw, soft_norm) if f]
                excluded = False
                for tok, tok_soft in exc_match:
                    if any(tok in field for field in fields):
                        excluded = True
                        break
                    if tok_soft and any(tok_soft in field for field in fields):
                        excluded = True
                        break
                if excluded:
                    debug_rows.append((raw, "matched exclude token"))
                    continue
            matched_token = None
            if inc_match:
                fields = [f for f in (low_raw, low_norm, soft_raw, soft_norm) if f]
                for tok, tok_soft in inc_match:
                    if any(tok in field for field in fields):
                        matched_token = tok
                        break
                    if tok_soft and any(tok_soft in field for field in fields):
                        matched_token = tok
                        break
            if inc and not matched_token:
                debug_rows.append((raw, "missing include token"))
                continue
            matches.append(
                {
                    "norm": norm,
                    "x": x,
                    "y": y,
                    "w": w,
                    "h": h,
                    "raw": raw,
                    "token": matched_token or "",
                    "soft": soft_raw,
                }
            )

        return matches, inc, exc, debug_rows

    def _prioritize_streitwert_matches(self, matches, inc):
        if not matches:
            return []
        if not inc:
            return matches

        ordered = []
        used_indices = set()

        for tok in inc:
            best_idx = None
            for idx, match in enumerate(matches):
                if idx in used_indices:
                    continue
                if match["token"] == tok:
                    best_idx = idx
                    break
            if best_idx is not None:
                ordered.append(matches[best_idx])
                used_indices.add(best_idx)

        for idx, match in enumerate(matches):
            if idx in used_indices:
                continue
            ordered.append(match)

        return ordered

    def _apply_ignore_top_doc_row(self, ordered, prefix=""):
        if (
            not ordered
            or not hasattr(self, "ignore_top_doc_row_var")
            or not bool(self.ignore_top_doc_row_var.get())
        ):
            return ordered

        top_match = min(
            ordered,
            key=lambda m: (m.get("y", 0), m.get("x", 0)),
        )

        remaining = [match for match in ordered if match is not top_match]
        self.log_print(
            f"{prefix}Ignoring top doc row match: {top_match.get('raw', '')}"
        )

        if not remaining:
            self.log_print(
                f"{prefix}No remaining Streitwert matches after ignoring top row."
            )

        return remaining

    def _doclist_abs_rect(self):
        if not self.current_rect:
            return None
        try:
            return rel_to_abs(self.current_rect, self.cfg["doclist_region"])
        except Exception:
            return None

    def _rechnungen_abs_rect(self):
        if not self.current_rect:
            return None
        region = self.cfg.get("rechnungen_region")
        if not (isinstance(region, (list, tuple)) and len(region) == 4):
            return None
        try:
            return rel_to_abs(self.current_rect, region)
        except Exception:
            return None

    def _prepare_ocr_variants(self, img, label=""):
        variants = []
        if img is None:
            return variants

        try:
            gray = img.convert("L")
        except Exception:
            try:
                gray = ImageOps.grayscale(img)
            except Exception:
                return [img]

        try:
            base_auto = ImageOps.autocontrast(gray)
        except Exception:
            base_auto = gray
        variants.append(base_auto)

        try:
            contrast_img = ImageEnhance.Contrast(gray).enhance(2.0)
            variants.append(ImageOps.autocontrast(contrast_img))
        except Exception:
            pass

        try:
            bright_img = ImageEnhance.Brightness(gray).enhance(1.2)
            variants.append(ImageOps.autocontrast(bright_img))
        except Exception:
            pass

        normalized_label = (label or "").strip().upper()

        if normalized_label == "GG":
            try:
                inverted = ImageOps.autocontrast(ImageOps.invert(gray))
                variants.append(inverted)
            except Exception:
                pass

            try:
                arr = np.array(gray, dtype=np.uint8)
                if arr.size:
                    if _HAS_CV2:
                        try:
                            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                            clahe_arr = clahe.apply(arr)
                            clahe_img = Image.fromarray(clahe_arr)
                            if clahe_img.mode != "L":
                                clahe_img = clahe_img.convert("L")
                            variants.append(ImageOps.autocontrast(clahe_img))
                        except Exception:
                            pass
                        blur = cv2.GaussianBlur(arr, (3, 3), 0)
                        _, thresh = cv2.threshold(
                            blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
                        )
                    else:
                        threshold = np.percentile(arr, 60)
                        thresh = (arr <= threshold).astype(np.uint8) * 255
                    thr_img = Image.fromarray(thresh.astype(np.uint8))
                    if thr_img.mode != "L":
                        thr_img = thr_img.convert("L")
                    try:
                        if ImageStat.Stat(thr_img).mean[0] > 127:
                            thr_img = ImageOps.invert(thr_img)
                    except Exception:
                        pass
                    try:
                        thr_img = ImageOps.autocontrast(thr_img)
                    except Exception:
                        pass
                    variants.append(thr_img)
            except Exception:
                pass

        unique = []
        seen = set()
        for candidate in variants:
            if candidate is None:
                continue
            try:
                hist = tuple(candidate.histogram())
                key = (candidate.mode, candidate.size, hist)
            except Exception:
                key = (candidate.mode, candidate.size)
            if key in seen:
                continue
            seen.add(key)
            unique.append(candidate)

        return unique or [base_auto]

    def _get_rechnungen_region_wait(self):
        wait_setting = self.cfg.get(
            "rechnungen_region_wait",
            DEFAULTS.get("rechnungen_region_wait", 0.0),
        )
        try:
            wait_seconds = float(wait_setting)
        except Exception:
            wait_seconds = float(DEFAULTS.get("rechnungen_region_wait", 0.0))
        return max(0.0, wait_seconds)

    def _wait_for_rechnungen_region(self, cfg_key, prefix="", label=""):
        wait_seconds = self._get_rechnungen_region_wait()
        if wait_seconds <= 0 or self._should_skip_manual_waits():
            return
        name = label or cfg_key.replace("_", " ").title()
        self.log_print(f"{prefix}Waiting {wait_seconds:.2f}s for {name} region.")
        time.sleep(wait_seconds)

    def _capture_named_region_preview_and_lines(self, cfg_key, prefix="", label=""):
        if not self.current_rect:
            self.log_print(
                f"{prefix}No active RDP rectangle. Connect before capturing."
            )
            return None, [], 1
        if not self._has(cfg_key):
            name = label or cfg_key.replace("_", " ").title()
            self.log_print(f"{prefix}{name} region is not configured.")
            return None, [], 1
        try:
            if cfg_key in {"rechnungen_region", "rechnungen_gg_region"}:
                self._wait_for_rechnungen_region(cfg_key, prefix=prefix, label=label)
            x, y, w, h = rel_to_abs(self.current_rect, self._get(cfg_key))
            scale = max(1, int(self.upscale_var.get() or 3))
            img = self._grab_region_color(x, y, w, h, upscale_x=scale)
        except Exception as exc:
            name = label or cfg_key.replace("_", " ").title()
            self.log_print(f"{prefix}Failed to capture {name} region: {exc}")
            return None, [], 1
        variants = self._prepare_ocr_variants(img, label=label)
        lines = []
        seen_lines = set()
        for variant in variants:
            try:
                df = do_ocr_data(
                    variant, lang=self.lang_var.get().strip() or "deu+eng", psm=6
                )
            except Exception:
                continue
            variant_lines = lines_from_tsv(df, scale=scale)
            for entry in variant_lines:
                if not (
                    isinstance(entry, (list, tuple))
                    and len(entry) == 5
                ):
                    continue
                x, y, w, h, text = entry
                text_key = " ".join(str(text).split()).lower()
                key = (
                    int(round(x / 4)) if x is not None else 0,
                    int(round(y / 4)) if y is not None else 0,
                    int(round(w / 4)) if w is not None else 0,
                    int(round(h / 4)) if h is not None else 0,
                    text_key,
                )
                if key in seen_lines:
                    continue
                seen_lines.add(key)
                lines.append(entry)
        lines.sort(key=lambda x: (x[1], x[0]))
        name = label or cfg_key.replace("_", " ").title()
        self.log_print(f"{prefix}{name} OCR lines: {len(lines)}.")
        return img, lines, scale

    def _capture_named_region_lines(self, cfg_key, prefix="", label=""):
        img, lines, _scale = self._capture_named_region_preview_and_lines(
            cfg_key, prefix=prefix, label=label
        )
        return lines

    def _capture_rechnungen_lines(self, prefix=""):
        return self._capture_named_region_lines(
            "rechnungen_region", prefix=prefix, label="Rechnungen"
        )

    def _capture_rechnungen_gg_lines(self, prefix=""):
        return self._capture_named_region_lines(
            "rechnungen_gg_region", prefix=prefix, label="GG"
        )

    def _merge_ocr_rows(self, lines):
        if not lines:
            return []

        def _safe_number(value, default=0.0):
            try:
                return float(value)
            except Exception:
                return float(default)

        heights = [
            _safe_number(h, 0.0)
            for _, _, _, h, _ in lines
            if isinstance(h, (int, float)) and h and _safe_number(h) > 0
        ]
        heights.sort()
        median_h = heights[len(heights) // 2] if heights else 12.0
        tolerance = max(4.0, median_h * 0.6)

        groups = []
        for entry in sorted(lines, key=lambda x: (x[1], x[0])):
            if not (isinstance(entry, (list, tuple)) and len(entry) == 5):
                continue
            x, y, w, h, text = entry
            raw_text = (text or "").strip()
            if not raw_text:
                continue
            x_val = _safe_number(x)
            y_val = _safe_number(y)
            w_val = max(_safe_number(w), 0.0)
            h_val = max(_safe_number(h, median_h), 0.0) or median_h
            center = y_val + h_val / 2.0

            target = None
            for group in groups:
                if abs(center - group["center"]) <= tolerance:
                    target = group
                    break

            if target is None:
                target = {
                    "items": [],
                    "min_x": x_val,
                    "min_y": y_val,
                    "max_x": x_val + max(w_val, 1.0),
                    "max_y": y_val + max(h_val, 1.0),
                    "center": center,
                }
                groups.append(target)
            else:
                target["min_x"] = min(target["min_x"], x_val)
                target["min_y"] = min(target["min_y"], y_val)
                target["max_x"] = max(target["max_x"], x_val + max(w_val, 1.0))
                target["max_y"] = max(target["max_y"], y_val + max(h_val, 1.0))
                target["center"] = (target["min_y"] + target["max_y"]) / 2.0

            target["items"].append(
                {
                    "x": x_val,
                    "y": y_val,
                    "w": w_val,
                    "h": h_val,
                    "text": raw_text,
                }
            )

        merged = []
        for group in groups:
            items_sorted = sorted(group["items"], key=lambda item: item["x"])
            pieces = [item["text"] for item in items_sorted if item.get("text")]
            if not pieces:
                continue
            combined = " ".join(pieces).strip()
            if not combined:
                continue
            min_x = int(round(group["min_x"]))
            min_y = int(round(group["min_y"]))
            width = int(round(max(1.0, group["max_x"] - group["min_x"])))
            height = int(round(max(1.0, group["max_y"] - group["min_y"])))
            tokens = [
                {
                    "x": int(round(item.get("x", 0.0))),
                    "y": int(round(item.get("y", 0.0))),
                    "w": int(round(max(item.get("w", 0.0), 1.0))),
                    "h": int(round(max(item.get("h", 0.0), 1.0))),
                    "text": item.get("text", ""),
                }
                for item in items_sorted
            ]
            merged.append(
                {
                    "x": min_x,
                    "y": min_y,
                    "w": width,
                    "h": height,
                    "text": combined,
                    "tokens": tokens,
                }
            )

        merged.sort(key=lambda item: (item.get("y", 0), item.get("x", 0)))
        return merged

    def _annotate_rechnungen_preview(self, img, entries, scale):
        if img is None:
            return None
        if not entries:
            return img
        try:
            preview = img.convert("RGB")
        except Exception:
            preview = img
        draw = ImageDraw.Draw(preview)
        thickness = max(1, int(scale // 2) or 1)
        for idx, entry in enumerate(entries, 1):
            if isinstance(entry.get("amount_box"), (list, tuple)) and len(entry["amount_box"]) == 4:
                base_x, base_y, base_w, base_h = entry["amount_box"]
            else:
                base_x = entry.get("x", 0)
                base_y = entry.get("y", 0)
                base_w = entry.get("w", 0)
                base_h = entry.get("h", 0)

            x = int(round(base_x * scale))
            y = int(round(base_y * scale))
            w = int(round(max(base_w, 1) * scale))
            h = int(round(max(base_h, 1) * scale))
            x1 = x + max(w, 1)
            y1 = y + max(h, 1)
            draw.rectangle([(x, y), (x1, y1)], outline="red", width=thickness)
            label = f"{idx}: {entry.get('amount', '')}".strip()
            if label:
                text_x = x + 2
                text_y = max(0, y - 14)
                try:
                    bbox = draw.textbbox((text_x, text_y), label)
                except Exception:
                    approx_w = max(32, 8 * len(label))
                    approx_h = 14
                    bbox = (text_x, text_y, text_x + approx_w, text_y + approx_h)
                draw.rectangle(
                    [(bbox[0] - 2, bbox[1] - 2), (bbox[2] + 2, bbox[3] + 2)],
                    fill=(255, 255, 255),
                )
                draw.text((text_x, text_y), label, fill="black")
        return preview

    def _select_rechnungen_amount_candidate(self, row, norm):
        tokens = row.get("tokens") or []
        row_x = row.get("x", 0)
        row_y = row.get("y", 0)
        row_w = row.get("w", 0)
        row_h = row.get("h", 0)
        candidates = []
        seen = set()

        for token in tokens:
            raw = token.get("text", "")
            if not raw:
                continue
            token_x = token.get("x", row_x)
            key_base = int(round(token_x))
            for amt in find_amount_candidates(raw):
                display = clean_amount_display(amt.get("display")) if amt else None
                if not display:
                    continue
                key = (display, key_base)
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(
                    {
                        "display": display,
                        "value": amt.get("value"),
                        "x": token_x,
                        "box": (
                            int(round(token.get("x", row_x))),
                            int(round(token.get("y", row_y))),
                            int(round(max(token.get("w", 0) or 1, 1))),
                            int(round(max(token.get("h", 0) or 1, 1))),
                        ),
                        "source": raw.strip(),
                    }
                )

        if not candidates:
            key_base = int(round(row_x))
            for amt in find_amount_candidates(norm):
                display = clean_amount_display(amt.get("display")) if amt else None
                if not display:
                    continue
                key = (display, key_base)
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(
                    {
                        "display": display,
                        "value": amt.get("value"),
                        "x": row_x,
                        "box": (
                            int(round(row_x)),
                            int(round(row_y)),
                            int(round(max(row_w, 1))),
                            int(round(max(row_h, 1))),
                        ),
                        "source": norm.strip(),
                    }
                )

        if not candidates:
            return {}

        zero = Decimal("0")

        def sort_key(candidate):
            value = candidate.get("value")
            positive = 0 if value is not None and value > zero else 1
            x_coord = candidate.get("x", row_x)
            magnitude = -value if value is not None else Decimal("0")
            return (positive, x_coord, magnitude)

        candidates.sort(key=sort_key)
        return candidates[0]

    def _extract_rechnungen_label_info(self, row, norm):
        tokens = row.get("tokens") or []
        row_x = row.get("x", 0)
        row_y = row.get("y", 0)
        row_w = max(row.get("w", 0), 1)
        label_tokens = []

        for token in tokens:
            raw = token.get("text", "")
            if not raw:
                continue
            normalized = normalize_line(raw)
            cleaned = re.sub(r"[^A-Z0-9]", "", normalized.upper())
            if not cleaned or not re.search(r"[A-Z]", cleaned):
                continue
            label_tokens.append(
                {
                    "raw": raw.strip(),
                    "clean": cleaned,
                    "x": token.get("x", row_x),
                    "y": token.get("y", row_y),
                    "w": int(round(max(token.get("w", 0), 1))),
                    "h": int(round(max(token.get("h", 0), 1))),
                }
            )

        if not label_tokens:
            return {"display": "", "normalized": normalize_gg_candidate(norm), "box": None}

        label_tokens.sort(key=lambda info: info.get("x", row_x))
        cutoff = row_x + row_w * 0.55
        tail = [info for info in label_tokens if info.get("x", row_x) >= cutoff]
        if not tail:
            tail = label_tokens[-3:]

        combined_raw = " ".join(info["raw"] for info in tail if info.get("raw"))
        combined_clean = "".join(info.get("clean", "") for info in tail)
        normalized = normalize_gg_candidate(combined_raw or combined_clean)

        min_x = min(info.get("x", row_x) for info in tail)
        min_y = min(info.get("y", row_y) for info in tail)
        max_x = max(info.get("x", row_x) + max(info.get("w", 1), 1) for info in tail)
        max_y = max(info.get("y", row_y) + max(info.get("h", 1), 1) for info in tail)
        box = (
            int(round(min_x)),
            int(round(min_y)),
            int(round(max_x - min_x)),
            int(round(max_y - min_y)),
        )

        return {
            "display": combined_raw.strip() or combined_clean,
            "normalized": normalized or normalize_gg_candidate(norm),
            "box": box,
        }

    def _parse_rechnungen_entries(self, lines, prefix=""):
        merged_lines = self._merge_ocr_rows(lines)
        entries = []
        skipped = []
        for row in merged_lines:
            raw = (row.get("text") or "").strip()
            if not raw:
                continue
            norm = normalize_line(raw)
            amount_info = self._select_rechnungen_amount_candidate(row, norm)
            amount = amount_info.get("display") if amount_info else None
            date_match = DATE_RE.search(norm) if norm else None
            if not amount or not date_match:
                if amount or date_match:
                    skipped.append((norm, "missing amount/date"))
                continue
            date_text = date_match.group(0)
            try:
                date_obj = datetime.strptime(date_text, "%d.%m.%Y")
            except ValueError:
                date_obj = None
            invoice_match = INVOICE_RE.search(norm)
            invoice = invoice_match.group(0) if invoice_match else ""
            label_info = self._extract_rechnungen_label_info(row, norm)
            label_display = label_info.get("display", "")
            label_normalized = label_info.get("normalized", "")
            entry = {
                "raw": raw,
                "norm": norm,
                "amount": clean_amount_display(amount) if amount else amount,
                "amount_box": amount_info.get("box") if amount_info else None,
                "amount_value": amount_info.get("value") if amount_info else None,
                "amount_source": amount_info.get("source") if amount_info else "",
                "date": date_text,
                "date_obj": date_obj,
                "invoice": invoice,
                "label": label_display,
                "label_normalized": label_normalized,
                "label_is_gg": (
                    is_gg_label(label_normalized) or is_gg_label(label_display)
                ),
                "label_box": label_info.get("box"),
                "x": row.get("x", 0),
                "y": row.get("y", 0),
                "w": row.get("w", 0),
                "h": row.get("h", 0),
            }
            entries.append(entry)
        entries.sort(
            key=lambda e: (
                e.get("date_obj") or datetime.min,
                e.get("y", 0),
                e.get("x", 0),
            )
        )
        for idx, entry in enumerate(entries, 1):
            detail = self._format_rechnungen_detail(entry)
            label_display = entry.get("label") or entry.get("label_normalized") or "-"
            amount_txt = entry.get("amount", "") or "(no amount)"
            bounds = entry.get("amount_box")
            bounds_txt = (
                f" | Bounds: ({bounds[0]}, {bounds[1]}, {bounds[2]}, {bounds[3]})"
                if isinstance(bounds, (list, tuple)) and len(bounds) == 4
                else ""
            )
            self.log_print(
                f"{prefix}Row {idx}: {amount_txt}{detail} | Label: {label_display}{bounds_txt}"
            )
        for norm, reason in skipped[:6]:
            self.log_print(f"{prefix}Skipped Rechnungen row '{norm}' ({reason}).")
        return entries

    def _is_gg_entry(self, entry):
        if not entry:
            return False
        if entry.get("label_is_gg"):
            return True

        label_norm = entry.get("label_normalized") or ""
        if is_gg_label(label_norm):
            return True

        label_display = entry.get("label") or ""
        if is_gg_label(label_display):
            return True

        raw_text = entry.get("raw") or ""
        if raw_text:
            tokens = re.split(r"[^A-Z0-9]+", normalize_line(raw_text).upper())
            for token in tokens:
                if not token or len(token) < 2:
                    continue
                if not re.search(r"[A-Z]", token):
                    continue
                if is_gg_label(token):
                    return True

        return False

    def _extract_rechnungen_gg_entries(self, prefix=""):
        wait_token = self._prepare_preview_wait()
        try:
            img, lines, scale = self._capture_named_region_preview_and_lines(
                "rechnungen_gg_region", prefix=prefix, label="GG"
            )
            if img is None and not lines:
                self._signal_preview_ready(wait_token=wait_token)
                return []

            entries = self._parse_rechnungen_entries(lines, prefix=prefix)
            gg_entries = [entry for entry in entries if self._is_gg_entry(entry)]

            log_lines = []
            if gg_entries:
                log_lines.append(
                    f"{prefix}Detected {len(gg_entries)} GG transaction(s)."
                )
            else:
                log_lines.append(f"{prefix}No GG transactions detected.")

            for idx, entry in enumerate(gg_entries, 1):
                amount = entry.get("amount", "") or "(no amount)"
                detail = self._format_rechnungen_detail(entry)
                label_display = (
                    entry.get("label")
                    or entry.get("label_normalized")
                    or "GG"
                )
                bounds = entry.get("amount_box")
                bounds_txt = (
                    f" | Bounds: ({bounds[0]}, {bounds[1]}, {bounds[2]}, {bounds[3]})"
                    if isinstance(bounds, (list, tuple)) and len(bounds) == 4
                    else ""
                )
                log_lines.append(
                    f"{prefix}GG #{idx}: {amount}{detail} | Label: {label_display}{bounds_txt}"
                )

            preview = self._annotate_rechnungen_preview(img, gg_entries, scale)
            if preview is not None:
                self.show_preview(preview, wait_token=wait_token)
            else:
                self._signal_preview_ready(wait_token=wait_token)

            for line in log_lines:
                self.log_print(line)

            return gg_entries
        except Exception:
            self._signal_preview_ready(wait_token=wait_token)
            raise

    def _summarize_rechnungen_entries(self, entries):
        def _copy(entry):
            if not entry:
                return {
                    "amount": "",
                    "date": "",
                    "invoice": "",
                    "raw": "",
                }
            return {
                "amount": entry.get("amount", ""),
                "date": entry.get("date", ""),
                "invoice": entry.get("invoice", ""),
                "raw": entry.get("raw", ""),
            }

        no_invoice = [e for e in entries if not e.get("invoice")]
        if no_invoice:
            no_invoice.sort(
                key=lambda e: (
                    e.get("date_obj") or datetime.min,
                    e.get("y", 0),
                    e.get("x", 0),
                )
            )
            total_entry = no_invoice[-1]
        else:
            total_entry = None

        invoice_entries = [e for e in entries if e.get("invoice")]
        invoice_entries.sort(
            key=lambda e: (
                e.get("date_obj") or datetime.min,
                e.get("y", 0),
                e.get("x", 0),
            )
        )

        court_entry = invoice_entries[-1] if invoice_entries else None
        gg_entry = invoice_entries[0] if len(invoice_entries) >= 2 else None

        summary = {
            "total": _copy(total_entry),
            "total_found": bool(total_entry),
            "court": _copy(court_entry),
            "court_found": bool(court_entry),
            "gg": (
                _copy(gg_entry)
                if gg_entry
                else {
                    "amount": "0",
                    "date": "",
                    "invoice": "",
                    "raw": "",
                }
            ),
            "gg_found": bool(gg_entry),
            "entries": [_copy(e) for e in entries],
        }
        if not summary["gg_found"]:
            if len(invoice_entries) == 1:
                summary["gg_missing_reason"] = "only one Rechnungen entry with invoice"
            elif not invoice_entries:
                summary["gg_missing_reason"] = "no Rechnungen entries with invoice"
        summary["invoice_entry_count"] = len(invoice_entries)
        summary["total_entry_count"] = len(no_invoice)
        return summary

    def _format_rechnungen_detail(self, entry):
        parts = []
        date = entry.get("date") if isinstance(entry, dict) else None
        invoice = entry.get("invoice") if isinstance(entry, dict) else None
        if date:
            parts.append(date)
        if invoice:
            parts.append(invoice)
        if not parts:
            return ""
        return f" ({' | '.join(parts)})"

    def _build_gg_summary_line(self, aktenzeichen, entries):
        label = aktenzeichen or "(unbekannt)"
        if not entries:
            return f"{label}: (no GG)"
        summary_parts = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            amount = entry.get("amount") or "(no amount)"
            detail = self._format_rechnungen_detail(entry)
            summary_parts.append(f"{amount}{detail}")
        if not summary_parts:
            return f"{label}: (no GG)"
        return f"{label}: {'; '.join(summary_parts)}"

    def _log_rechnungen_summary(self, prefix, summary):
        if not summary:
            self.log_print(f"{prefix}-Total Fees: (not found)")
            self.log_print(f"{prefix}-Received Court Fees: (not found)")
            self.log_print(f"{prefix}-Received GG: 0")
            return

        total_txt = summary.get("total", {}).get("amount", "") or "(not found)"
        self.log_print(f"{prefix}-Total Fees: {total_txt}")

        court_entry = summary.get("court", {})
        court_amt = court_entry.get("amount", "")
        if court_amt:
            detail = self._format_rechnungen_detail(court_entry)
            self.log_print(f"{prefix}-Received Court Fees: {court_amt}{detail}")
        else:
            self.log_print(f"{prefix}-Received Court Fees: (not found)")

        gg_entry = summary.get("gg", {})
        if summary.get("gg_found"):
            detail = self._format_rechnungen_detail(gg_entry)
            self.log_print(
                f"{prefix}-Received GG: {gg_entry.get('amount', '')}{detail}"
            )
        else:
            gg_amt = gg_entry.get("amount", "0") or "0"
            self.log_print(f"{prefix}-Received GG: {gg_amt}")
            reason = summary.get("gg_missing_reason")
            if reason:
                self.log_print(f"{prefix}  ↳ Reason: {reason}.")

    def _build_rechnungen_result_row(self, aktenzeichen, summary):
        if not summary:
            summary = {
                "total": {"amount": "", "date": "", "invoice": ""},
                "court": {"amount": "", "date": "", "invoice": ""},
                "gg": {"amount": "0", "date": "", "invoice": ""},
            }
        total = summary.get("total", {})
        court = summary.get("court", {})
        gg = summary.get("gg", {})
        return {
            "aktenzeichen": aktenzeichen,
            "total_fees_amount": total.get("amount", ""),
            "total_fees_date": total.get("date", ""),
            "total_fees_invoice": total.get("invoice", ""),
            "received_court_amount": court.get("amount", ""),
            "received_court_date": court.get("date", ""),
            "received_court_invoice": court.get("invoice", ""),
            "received_gg_amount": gg.get("amount", ""),
            "received_gg_date": gg.get("date", ""),
            "received_gg_invoice": gg.get("invoice", ""),
        }

    def _extract_rechnungen_summary(self, prefix=""):
        lines = self._capture_rechnungen_lines(prefix=prefix)
        if not lines:
            return None
        entries = self._parse_rechnungen_entries(lines, prefix=prefix)
        if not entries:
            return self._summarize_rechnungen_entries(entries)
        return self._summarize_rechnungen_entries(entries)

    def _select_doclist_entry(self, match, doc_rect, focus_first=False, prefix=""):
        if not match or not doc_rect:
            self.log_print(f"{prefix}No match/doc_rect provided for selection.")
            return False
        rx, ry, rw, rh = doc_rect
        if focus_first:
            focus_x = rx + max(5, rw // 40)
            focus_y = ry + max(5, rh // 40)
            self.log_print(f"{prefix}Focusing doc list at ({focus_x}, {focus_y}).")
            pyautogui.click(focus_x, focus_y)
            time.sleep(0.2)
        local_x = match["x"] + max(
            12, min(match["w"] // 2 if match["w"] else 20, max(match["w"] - 10, 12))
        )
        local_y = match["y"] + max(
            10, min(match["h"] // 2 if match["h"] else 20, max(match["h"] - 10, 10))
        )
        local_x = max(5, local_x)
        local_y = max(5, local_y)
        click_x = rx + min(local_x, max(rw - 5, 5))
        click_y = ry + min(local_y, max(rh - 5, 5))
        self.log_print(
            f"{prefix}Moving to row '{match['raw']}' at ({click_x}, {click_y}) size ({match['w']}x{match['h']})."
        )
        pyautogui.moveTo(click_x, click_y)
        pyautogui.click(click_x, click_y)
        time.sleep(0.2)
        self.log_print(f"{prefix}Clicked row '{match['raw']}'.")
        return True

    def _click_view_button(self, prefix=""):
        if not self.current_rect:
            return False
        point = self.cfg.get("doc_view_point")
        if not (isinstance(point, (list, tuple)) and len(point) == 2):
            msg = "View button point is not configured. Please calibrate it."
            self.log_print(f"{prefix}{msg}" if prefix else msg)
            return False
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass
        vx, vy = rel_to_abs(self.current_rect, point)
        self.log_print(f"{prefix}Clicking View button at ({vx}, {vy}).")
        pyautogui.moveTo(vx, vy)
        pyautogui.click(vx, vy)
        time.sleep(0.1)
        return True

    def _find_overlay_entry(self, lines):
        for x, y, w, h, raw in lines:
            if raw is None:
                continue
            norm = normalize_line(raw)
            candidates = [
                str(raw).strip().lower(),
                norm.lower() if norm else "",
            ]
            ascii_candidates = []
            for cand in candidates:
                if cand:
                    ascii_candidates.append(
                        unicodedata.normalize("NFKD", cand)
                        .encode("ascii", "ignore")
                        .decode("ascii")
                    )
            candidates.extend(ascii_candidates)
            for candidate in candidates:
                if not candidate:
                    continue
                for pattern in DOC_LOADING_PATTERNS:
                    if pattern in candidate:
                        return {
                            "raw": str(raw),
                            "norm": norm,
                            "x": int(x),
                            "y": int(y),
                            "w": int(w),
                            "h": int(h),
                        }
        return None

    def _detect_overlay_in_rel_box(self, rel_box):
        # Existing code
        if not self.current_rect or not rel_box:
            return None

        rx, ry, rw, rh = rel_to_abs(self.current_rect, rel_box)
        img, scale = _grab_region_color_generic(
            self.current_rect, rel_box, self.upscale_var.get()
        )
        df = do_ocr_data(img, lang=self.lang_var.get().strip() or "deu+eng", psm=6)
        lines = lines_from_tsv(df, scale=scale)
        overlay = self._find_overlay_entry(lines)
        if overlay:
            entry = overlay.copy()
            entry["abs_x"] = rx + overlay["x"]
            entry["abs_y"] = ry + overlay["y"]
            entry["abs_w"] = overlay["w"]
            entry["abs_h"] = overlay["h"]
            # <<< LOG >>>
            self.log_print(
                f"[OVERLAY DETECTED] {rel_box} overlay found: '{entry.get('norm', '')}'"
            )
            return entry
        else:
            # <<< LOG >>>
            lines_text = [l[4] for l in lines if len(l) >= 5]
            if any(lines_text):
                self.log_print(
                    f"[OVERLAY CHECK] {rel_box} no overlay match: '{', '.join(t for t in lines_text if t)}'"
                )
        return None

    def _wait_for_doclist_ready(self, prefix="", timeout=12.0, reason=""):
        doc_rect = self._doclist_abs_rect()
        if not doc_rect:
            return
        suffix = f" ({reason})" if reason else ""
        start = time.time()
        notified = False
        last_log = 0.0
        while True:
            overlay = self._detect_overlay_in_rel_box(self.cfg["doclist_region"])
            if not overlay:
                if notified:
                    self.log_print(f"{prefix}Document list overlay cleared{suffix}.")
                    # <<< LOG >>>
                    self.log_print(
                        "[OVERLAY CLEARED] Document list ready, no waiting box detected."
                    )
                return
            now = time.time()
            desc = (
                overlay.get("norm")
                or normalize_line(overlay.get("raw"))
                or overlay.get("raw")
                or "(overlay text not recognized)"
            )
            coords = (
                overlay.get("abs_x", 0),
                overlay.get("abs_y", 0),
                overlay.get("abs_w", 0),
                overlay.get("abs_h", 0),
            )
            if (not notified) or (now - last_log >= 1.5):
                self.log_print(
                    f"{prefix}Document list overlay detected{suffix}: '{desc}' at ({coords[0]}, {coords[1]}, {coords[2]}x{coords[3]}). Waiting..."
                )
                last_log = now
            notified = True
            if time.time() - start > timeout:
                self.log_print(
                    f"{prefix}Timeout waiting for document list overlay to clear{suffix}. Continuing."
                )
                return
            time.sleep(0.5)

    def _search_overlay_rel_box(self):
        if not self.current_rect:
            return None
        point = self.cfg.get("search_point")
        if not (isinstance(point, (list, tuple)) and len(point) == 2):
            return None
        left, top, right, bottom = self.current_rect
        width = max(1, right - left)
        height = max(1, bottom - top)
        target_w = min(420, width)
        target_h = min(220, height)
        rel_w = target_w / width
        rel_h = target_h / height
        rel_left = max(0.0, min(point[0] - rel_w / 2, 1 - rel_w))
        rel_top = max(0.0, min(point[1] - rel_h / 2, 1 - rel_h))
        return [rel_left, rel_top, rel_w, rel_h]

    def _wait_for_doc_search_ready(self, prefix="", timeout=10.0, reason=""):
        rel_box = self._search_overlay_rel_box()
        if not rel_box:
            return
        suffix = f" ({reason})" if reason else ""
        start = time.time()
        notified = False
        last_log = 0.0
        while True:
            overlay = self._detect_overlay_in_rel_box(rel_box)
            if not overlay:
                if notified:
                    self.log_print(f"{prefix}Deal search overlay cleared{suffix}.")
                    # <<< LOG >>>
                    self.log_print(
                        "[OVERLAY CLEARED] Document search ready, no waiting box detected."
                    )
                return
            now = time.time()
            desc = (
                overlay.get("norm")
                or normalize_line(overlay.get("raw"))
                or overlay.get("raw")
                or "(overlay text not recognized)"
            )
            coords = (
                overlay.get("abs_x", 0),
                overlay.get("abs_y", 0),
                overlay.get("abs_w", 0),
                overlay.get("abs_h", 0),
            )
            if (not notified) or (now - last_log >= 1.5):
                self.log_print(
                    f"{prefix}Deal search overlay detected{suffix}: '{desc}' at ({coords[0]}, {coords[1]}, {coords[2]}x{coords[3]}). Waiting..."
                )
                last_log = now
            notified = True
            if time.time() - start > timeout:
                self.log_print(
                    f"{prefix}Timeout waiting for deal search overlay to clear{suffix}. Continuing."
                )
                return
            time.sleep(0.4)

    def _wait_for_pdf_ready(self, prefix="", timeout=12.0, reason=""):
        if not self.current_rect or "pdf_text_region" not in self.cfg:
            return
        suffix = f" ({reason})" if reason else ""
        start = time.time()
        notified = False
        last_log = 0.0
        while True:
            overlays = []
            pdf_box = self.cfg.get("pdf_text_region")
            if pdf_box:
                overlay_pdf = self._detect_overlay_in_rel_box(pdf_box)
                if overlay_pdf:
                    overlay_pdf["area"] = "PDF view"
                    overlays.append(overlay_pdf)
            doc_box = self.cfg.get("doclist_region")
            if doc_box:
                overlay_doc = self._detect_overlay_in_rel_box(doc_box)
                if overlay_doc:
                    overlay_doc["area"] = "Document list"
                    overlays.append(overlay_doc)
            search_box = self._search_overlay_rel_box()
            if search_box:
                overlay_search = self._detect_overlay_in_rel_box(search_box)
                if overlay_search:
                    overlay_search["area"] = "Document search"
                    overlays.append(overlay_search)
            if not overlays:
                if notified:
                    self.log_print(f"{prefix}All PDF overlays cleared{suffix}.")
                    # <<< LOG >>>
                    self.log_print(
                        "[OVERLAY CLEARED] PDF ready, no waiting box detected."
                    )
                return
            now = time.time()
            if (not notified) or (now - last_log >= 1.5):
                for entry in overlays:
                    desc = (
                        entry.get("norm")
                        or normalize_line(entry.get("raw"))
                        or entry.get("raw")
                        or "(overlay text not recognized)"
                    )
                    coords = (
                        entry.get("abs_x", 0),
                        entry.get("abs_y", 0),
                        entry.get("abs_w", 0),
                        entry.get("abs_h", 0),
                    )
                    area = entry.get("area", "Overlay")
                    self.log_print(
                        f"{prefix}{area} overlay detected{suffix}: '{desc}' at ({coords[0]}, {coords[1]}, {coords[2]}x{coords[3]}). Waiting..."
                    )
                last_log = now
            notified = True
            if now - start > timeout:
                self.log_print(
                    f"{prefix}Timeout waiting for PDF overlays to clear{suffix}. Continuing."
                )
                return
            time.sleep(0.4)

    def _type_pdf_search(self, query, prefix="", press_enter=True):
        if not self.current_rect:
            return False
        point = self.cfg.get("pdf_search_point")
        if not (isinstance(point, (list, tuple)) and len(point) == 2):
            msg = "PDF search point is not configured. Please calibrate it."
            self.log_print(f"{prefix}{msg}" if prefix else msg)
            return False
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass
        sx, sy = rel_to_abs(self.current_rect, point)
        self.log_print(
            f"{prefix}Clicking PDF search box at ({sx}, {sy}) and typing '{query}'."
        )
        pyautogui.click(sx, sy)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
        pyautogui.typewrite(query or "", interval=float(self.type_var.get() or 0.02))
        if press_enter:
            self.log_print(f"{prefix}Pressing Enter to run PDF search.")
            pyautogui.press("enter")
        return True

    def _click_pdf_result_button(self, prefix="", which="primary"):
        if not self.current_rect:
            return False
        mapping = {
            "primary": ("pdf_hits_point", "PDF result button"),
            "secondary": ("pdf_hits_second_point", "secondary PDF result button"),
            "tertiary": ("pdf_hits_third_point", "tertiary PDF result button"),
        }
        key, label = mapping.get(which, mapping["primary"])
        point = self.cfg.get(key)
        if not (isinstance(point, (list, tuple)) and len(point) == 2):
            msg = f"{label.capitalize()} is not configured. Please calibrate it."
            self.log_print(f"{prefix}{msg}" if prefix else msg)
            return False
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass
        hx, hy = rel_to_abs(self.current_rect, point)
        self.log_print(f"{prefix}Clicking {label} at ({hx}, {hy}).")
        pyautogui.moveTo(hx, hy)
        pyautogui.click(hx, hy)
        wait_seconds = float(self.hitwait_var.get() or 1.0)
        if wait_seconds > 0 and not self._should_skip_manual_waits():
            time.sleep(wait_seconds)
        return True

    def _click_pdf_close_button(self, prefix=""):
        if not self.current_rect:
            return False
        point = self.cfg.get("pdf_close_point")
        if not (isinstance(point, (list, tuple)) and len(point) == 2):
            msg = "PDF close button point is not configured. Please calibrate it."
            self.log_print(f"{prefix}{msg}" if prefix else msg)
            return False
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass
        cx, cy = rel_to_abs(self.current_rect, point)
        self.log_print(f"{prefix}Clicking PDF close button at ({cx}, {cy}).")
        pyautogui.moveTo(cx, cy)
        pyautogui.click(cx, cy)
        time.sleep(0.3)
        return True

    def _type_doclist_query(self, query, press_enter=True, prefix=""):
        if not self.current_rect:
            return False
        try:
            sx, sy = rel_to_abs(self.current_rect, self.cfg["search_point"])
        except Exception:
            self.log_print(
                "Doc list search point is not configured. Please calibrate the search point."
            )
            return False

        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        self.log_print(
            f"{prefix}Clicking search box at ({sx}, {sy}) and typing '{query}'."
        )
        pyautogui.click(sx, sy)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
        pyautogui.typewrite(query or "", interval=float(self.type_var.get() or 0.02))
        if press_enter:
            self.log_print(f"{prefix}Pressing Enter after typing query.")
            pyautogui.press("enter")
        return True

    def _close_active_pdf(self, prefix=""):
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            return
        if self._click_pdf_close_button(prefix=prefix):
            self.log_print(f"{prefix}Requested PDF close via window button.")
            time.sleep(0.6)
            return
        self.log_print(
            f"{prefix}PDF close button not configured; using keyboard shortcuts."
        )
        pyautogui.hotkey("ctrl", "w")
        time.sleep(0.4)
        pyautogui.hotkey("ctrl", "f4")
        time.sleep(0.2)

    def _process_open_pdf(
        self, prefix="", search_term=None, retype=False, log_label=None
    ):
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass

        self._wait_for_pdf_ready(prefix=prefix, reason="after opening PDF")

        term = search_term
        if term is None:
            term = self.streitwort_var.get().strip() or "Streitwert"

        if retype and term:
            if not self._type_pdf_search(term, prefix=prefix):
                self.log_print(
                    f"{prefix}Unable to type PDF search term; continuing without re-search."
                )
            else:
                wait_seconds = float(self.hitwait_var.get() or 1.0)
                if wait_seconds > 0 and not self._should_skip_manual_waits():
                    time.sleep(wait_seconds)
                self._wait_for_pdf_ready(prefix=prefix, reason="after PDF search")

        try:
            extra_wait = float(self.cfg.get("pdf_view_extra_wait", 2.0))
        except Exception:
            extra_wait = 2.0

        clicked_hits = self._click_pdf_result_button(prefix=prefix)
        if clicked_hits:
            if extra_wait > 0:
                self.log_print(
                    f"{prefix}Waiting {extra_wait:.1f}s after PDF results click before checking overlays."
                )
                time.sleep(extra_wait)
            primary_reason = "after PDF results click"
        else:
            self.log_print(
                f"{prefix}Skipped PDF results click; proceeding directly to page OCR."
            )
            primary_reason = "before page OCR"
        keyword_candidates = build_streitwert_keywords(term)

        def extract_from_current_page(reason_label, attempt_label="primary"):
            self._wait_for_pdf_ready(prefix=prefix, reason=reason_label)
            x, y, w, h = rel_to_abs(self.current_rect, self._get("pdf_text_region"))
            page_img = self._grab_region_color(
                x, y, w, h, upscale_x=self.upscale_var.get()
            )
            page_scale = max(1, int(self.upscale_var.get() or 3))
            dft = do_ocr_data(
                page_img, lang=self.lang_var.get().strip() or "deu+eng", psm=6
            )
            lines_pg = lines_from_tsv(dft, scale=page_scale)
            self._append_ocr_log(
                log_label,
                f"pdf_{attempt_label}",
                lines_pg,
                prefix=prefix,
                keywords=keyword_candidates,
            )
            self.log_print(
                f"{prefix}Page OCR lines captured ({attempt_label}): {len(lines_pg)}. Extracting amount."
            )
            amount_raw, amount_line = extract_amount_from_lines(
                lines_pg,
                keyword=keyword_candidates or None,
                min_value=STREITWERT_MIN_AMOUNT,
            )
            amount_clean = clean_amount_display(amount_raw) if amount_raw else None
            if amount_clean and prefix:
                suffix = f" [{attempt_label}]" if attempt_label else ""
                self.log_print(
                    f"{prefix}Matched Streitwert line{suffix}: {amount_line or '(context unavailable)'}"
                )
            return amount_clean

        amount = extract_from_current_page(primary_reason, attempt_label="primary")
        if amount:
            return amount

        second_point = self.cfg.get("pdf_hits_second_point")
        if isinstance(second_point, (list, tuple)) and len(second_point) == 2:
            self.log_print(
                f"{prefix}Primary PDF result yielded no amount. Trying secondary result."
            )
            if self._click_pdf_result_button(prefix=prefix, which="secondary"):
                if extra_wait > 0:
                    self.log_print(
                        f"{prefix}Waiting {extra_wait:.1f}s after secondary PDF result click."
                    )
                    time.sleep(extra_wait)
                amount = extract_from_current_page(
                    "after secondary PDF result click", attempt_label="secondary"
                )
                if amount:
                    return amount
            else:
                self.log_print(
                    f"{prefix}Secondary PDF result button not available; skipping fallback."
                )

        third_point = self.cfg.get("pdf_hits_third_point")
        if isinstance(third_point, (list, tuple)) and len(third_point) == 2:
            self.log_print(
                f"{prefix}No Streitwert from earlier attempts. Trying third PDF result."
            )
            if self._click_pdf_result_button(prefix=prefix, which="tertiary"):
                if extra_wait > 0:
                    self.log_print(
                        f"{prefix}Waiting {extra_wait:.1f}s after third PDF result click."
                    )
                    time.sleep(extra_wait)
                amount = extract_from_current_page(
                    "after third PDF result click", attempt_label="tertiary"
                )
                if amount:
                    return amount
            else:
                self.log_print(
                    f"{prefix}Tertiary PDF result button not available; skipping fallback."
                )

        return None

    def _gather_aktenzeichen(self):
        try:
            df = pd.read_excel(
                self.cfg["excel_path"], sheet_name=self.cfg["excel_sheet"]
            )
        except Exception as exc:
            self.log_print(f"Failed to open Excel file: {exc}")
            return []

        start_cell = (self.start_cell_var.get() or "").strip()
        max_rows = int(self.max_rows_var.get() or "0")

        if start_cell:
            m = re.match(r"^\s*([A-Za-z]+)\s*([0-9]+)\s*$", start_cell)
            if not m:
                self.log_print(
                    f"Invalid start cell '{start_cell}'. Use spreadsheet format like 'B2'."
                )
                return []
            col_letters, row_num = m.group(1).upper(), int(m.group(2))
            col_idx = 0
            for ch in col_letters:
                col_idx = col_idx * 26 + (ord(ch) - 64)
            col_idx -= 1
            rows = df.iloc[max(row_num - 2, 0) :]
        else:
            column = self.cfg["input_column"]
            if column not in df.columns:
                self.log_print(
                    f"Column '{column}' not found in the Excel sheet for Streitwert scan."
                )
                return []
            col_idx = df.columns.get_loc(column)
            rows = df

        if max_rows > 0:
            rows = rows.head(max_rows)

        queries = []
        for _, row in rows.iterrows():
            q = str(row.iloc[col_idx]).strip()
            if q and q.lower() != "nan":
                queries.append((q, row.to_dict()))

        if not queries:
            self.log_print(
                "No Aktenzeichen values were found in the configured Excel sheet."
            )
        return queries

    def run_streitwert(self, include_rechnungen=False):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()

            doc_rect = self._doclist_abs_rect()
            if not doc_rect:
                self.log_print(
                    "Doc list region is not configured. Please re-run calibration."
                )
                return

            queries = self._gather_aktenzeichen()
            self.clear_simple_log()
            self._reset_ocr_log_state()
            if not queries:
                return

            skip_waits = self._should_skip_manual_waits()
            list_wait = (
                0.0 if skip_waits else float(self.cfg.get("post_search_wait", 1.2))
            )
            doc_wait = 0.0 if skip_waits else float(self.docwait_var.get() or 1.2)
            results = []
            rechnungen_results = []
            total = len(queries)
            for idx, (aktenzeichen, _row) in enumerate(queries, 1):
                prefix = f"[{idx}/{total}] "
                self.log_print(
                    f"{prefix}Searching doc list for Aktenzeichen: {aktenzeichen}"
                )
                if not self._type_doclist_query(aktenzeichen, prefix=prefix):
                    self.log_print(
                        f"{prefix}Unable to type Aktenzeichen. Skipping entry."
                    )
                    continue
                if list_wait > 0:
                    time.sleep(list_wait)
                self.log_print(
                    f"{prefix}Typed '{aktenzeichen}' into the document search box."
                )
                self._wait_for_doc_search_ready(
                    prefix=prefix, reason="after Aktenzeichen search"
                )
                self._wait_for_doclist_ready(
                    prefix=prefix, reason="after Aktenzeichen search"
                )

                rechn_summary = None
                if include_rechnungen:
                    rechn_summary = self._extract_rechnungen_summary(prefix=prefix)
                    if rechn_summary is None:
                        self.log_print(
                            f"{prefix}Rechnungen capture returned no data; storing defaults."
                        )
                        rechn_summary = self._summarize_rechnungen_entries([])
                    else:
                        self._log_rechnungen_summary(prefix, rechn_summary)
                    rechn_row = self._build_rechnungen_result_row(
                        aktenzeichen, rechn_summary
                    )
                    inst_info = self.detect_instance()
                    rechn_row["instance_detected"] = (inst_info or {}).get("instance")
                    rechnungen_results.append(rechn_row)

                term = self.streitwort_var.get().strip() or "Streitwert"
                self._wait_for_doc_search_ready(
                    prefix=prefix, reason="before PDF search"
                )
                if not self._type_pdf_search(term, prefix=prefix):
                    self.log_print(
                        f"{prefix}Unable to type Streitwert term in the PDF search box."
                    )
                    continue
                self.log_print(f"{prefix}Typed '{term}' into the PDF search box.")
                if list_wait > 0:
                    time.sleep(list_wait)
                self._wait_for_doc_search_ready(
                    prefix=prefix, reason="after PDF search"
                )
                self._wait_for_doclist_ready(prefix=prefix, reason="after PDF search")

                rx, ry, rw, rh = doc_rect
                focus_x = rx + max(5, rw // 40)
                focus_y = ry + max(5, rh // 40)
                self.log_print(
                    f"{prefix}Clicking doc list to ensure focus at ({focus_x}, {focus_y})."
                )
                pyautogui.click(focus_x, focus_y)
                time.sleep(0.2)

                x, y, w, h = rel_to_abs(self.current_rect, self._get("doclist_region"))
                doc_img = self._grab_region_color(
                    x, y, w, h, upscale_x=self.upscale_var.get()
                )
                doc_scale = max(1, int(self.upscale_var.get() or 3))
                df = do_ocr_data(
                    doc_img,
                    lang=self.lang_var.get().strip() or "deu+eng",
                    psm=6,
                )
                lines = lines_from_tsv(df, scale=doc_scale)
                matches, inc, exc, debug_rows = self._filter_streitwert_rows(lines)
                ordered = self._prioritize_streitwert_matches(matches, inc)
                ordered = self._apply_ignore_top_doc_row(ordered, prefix=prefix)

                if not ordered:
                    self._append_ocr_log(
                        f"{aktenzeichen}_doclist_nomatch",
                        "doclist",
                        lines,
                        prefix=prefix,
                    )
                    reason = ", ".join(f"{r}: {raw}" for raw, r in debug_rows[:4])
                    if not reason:
                        sample = ", ".join((txt or "").strip() for *_, txt in lines[:4])
                        reason = sample or "no OCR rows"
                    self.log_print(
                        f"{prefix}No matching rows for '{aktenzeichen}'. Details: {reason}"
                    )
                    continue

                first = ordered[0]
                tag = first.get("token") or "any"
                log_label = (
                    f"{aktenzeichen}_{first.get('raw', '')}"
                    if aktenzeichen
                    else first.get("raw", "")
                )
                self._append_ocr_log(
                    log_label,
                    "doclist",
                    lines,
                    prefix=prefix,
                )
                preview = ", ".join(
                    f"{m['token'] or 'any'} → {m['raw']}" for m in ordered[:3]
                )
                self.log_print(
                    f"{prefix}Selecting {tag} match: {first['raw']} | candidates: {preview}"
                )

                if not self._select_doclist_entry(
                    first,
                    doc_rect,
                    focus_first=True,
                    prefix=prefix,
                ):
                    self.log_print(
                        f"{prefix}Unable to activate doc row: {first.get('raw','')}"
                    )
                    continue

                if not self._click_view_button(prefix=prefix):
                    self.log_print(
                        f"{prefix}Skipping entry because the View button click failed."
                    )
                    continue

                self.log_print(f"{prefix}Clicked View button for the selected row.")

                if doc_wait > 0:
                    time.sleep(doc_wait)
                amount = self._process_open_pdf(
                    prefix=prefix,
                    log_label=log_label,
                )
                inst_info = self.detect_instance()
                rec = {
                    "aktenzeichen": aktenzeichen,
                    "row_text": first["norm"],
                    "amount": amount or "",
                    "instance_detected": (inst_info or {}).get("instance"),
                }
                results.append(rec)
                self.simple_log_print(f"{aktenzeichen}: {amount or '(none)'}")
                self.log_print(
                    f"{prefix}{aktenzeichen} / {first['norm']} → {amount or '(none)'}"
                )

                self._close_active_pdf(prefix=prefix)
                time.sleep(0.5)

            if results:
                pd.DataFrame(results).to_csv(
                    self.streit_csv_var.get(), index=False, encoding="utf-8-sig"
                )
                self.log_print(
                    f"Done. Saved Streitwert results to {self.streit_csv_var.get()}"
                )
            else:
                self.log_print(
                    "No Streitwert results were collected from the Excel list."
                )

            if include_rechnungen:
                if rechnungen_results:
                    pd.DataFrame(rechnungen_results).to_csv(
                        self.rechnungen_csv_var.get(),
                        index=False,
                        encoding="utf-8-sig",
                    )
                    self.log_print(
                        f"Done. Saved Rechnungen results to {self.rechnungen_csv_var.get()}"
                    )
                else:
                    self.log_print(
                        "No Rechnungen values were captured from the Excel list."
                    )

        except Exception as e:
            self.log_print("ERROR: " + repr(e))

    def test_streitwert_setup(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()

            doc_rect = self._doclist_abs_rect()
            if not doc_rect:
                self.log_print("[Test] Doc list region not configured.")
                return

            term = self.streitwort_var.get() or "Streitwert"
            skip_waits = self._should_skip_manual_waits()
            list_wait = (
                0.0 if skip_waits else float(self.cfg.get("post_search_wait", 1.2))
            )
            doc_wait = 0.0 if skip_waits else float(self.docwait_var.get() or 1.2)
            self._reset_ocr_log_state()
            if not self._type_pdf_search(term, prefix="[Test] "):
                self.log_print("[Test] Unable to type the Streitwert search term.")
                return
            self.log_print(f"[Test] Typed '{term}' into the PDF search box.")
            inst_info = self.detect_instance(prefix="[Test] ")
            if list_wait > 0:
                time.sleep(list_wait)
            self._wait_for_doc_search_ready(prefix="[Test] ", reason="after PDF search")
            self._wait_for_doclist_ready(prefix="[Test] ", reason="after PDF search")

            rx, ry, rw, rh = doc_rect
            focus_x = rx + max(5, rw // 40)
            focus_y = ry + max(5, rh // 40)
            self.log_print(
                f"[Test] Clicking doc list to ensure focus at ({focus_x}, {focus_y})."
            )
            pyautogui.click(focus_x, focus_y)
            time.sleep(0.2)

            doc_img, doc_scale = _grab_region_color_generic(
                self.current_rect,
                self.cfg["doclist_region"],
                self.upscale_var.get(),
            )
            df = do_ocr_data(
                doc_img,
                lang=self.lang_var.get().strip() or "deu+eng",
                psm=6,
            )
            lines = lines_from_tsv(df, scale=doc_scale)
            matches, inc, exc, debug_rows = self._filter_streitwert_rows(lines)
            ordered = self._prioritize_streitwert_matches(matches, inc)
            ordered = self._apply_ignore_top_doc_row(ordered, prefix="[Test] ")

            self.log_print(
                f"[Test] Doc list OCR lines: {len(lines)} | includes: {inc or ['(none)']} | excludes: {exc or ['(none)']}"
            )
            if not ordered:
                self._append_ocr_log(
                    "TEST_doclist_nomatch",
                    "doclist",
                    lines,
                    prefix="[Test] ",
                )
                preview = debug_rows[:5] or [(raw, "") for *_, raw in lines[:5]]
                for raw, reason in preview:
                    desc = f"  {reason or 'OCR'} → {raw}"
                    self.log_print(desc)
                self.log_print(
                    "[Test] No rows matched the include tokens after typing 'Streitwert'."
                )
                return

            first = ordered[0]
            tag = first.get("token") or "any"
            log_label = f"TEST_{first.get('raw', '')}"
            self._append_ocr_log(
                log_label,
                "doclist",
                lines,
                prefix="[Test] ",
            )
            if self._select_doclist_entry(
                first, doc_rect, focus_first=True, prefix="[Test] "
            ):
                self.log_print(
                    f"[Test] Selected first matching row ({tag}): {first['raw']}"
                )
            else:
                self.log_print("[Test] Failed to select the first matching row.")
                return

            if not self._click_view_button(prefix="[Test] "):
                self.log_print("[Test] Unable to click the View button.")
                return

            self.log_print("[Test] Clicked View button to open the PDF.")
            if doc_wait > 0:
                time.sleep(doc_wait)
            amount = self._process_open_pdf(
                prefix="[Test] ",
                search_term=term or "Streitwert",
                log_label=log_label,
            )
            self.log_print(f"[Test] Extracted Streitwert amount: {amount or '(none)'}")

            self._close_active_pdf(prefix="[Test] ")
            time.sleep(0.5)
            self.log_print("[Test] Closed PDF after verification.")

            self.log_print("[Test] Streitwert setup check finished.")

        except Exception as e:
            self.log_print("ERROR during Streitwert test: " + repr(e))

    # ---------- Utilities ----------
    def _prepare_preview_wait(self):
        evt = getattr(self, "_preview_ready_event", None)
        if evt:
            evt.clear()
        last = getattr(self, "_preview_last_token", 0)
        target = getattr(self, "_preview_target_token", 0)
        token = max(last, target) + 1
        self._preview_target_token = token
        return token

    def _signal_preview_ready(self, wait_token=None):
        last = getattr(self, "_preview_last_token", 0)
        target = getattr(self, "_preview_target_token", 0)
        if wait_token is None:
            wait_token = max(last + 1, target)
        else:
            wait_token = max(wait_token, last)
        self._preview_last_token = wait_token
        evt = getattr(self, "_preview_ready_event", None)
        if evt:
            evt.set()

    def _wait_for_preview_ready(self, timeout=1.5):
        target = getattr(self, "_preview_target_token", 0)
        if target <= 0:
            if timeout and timeout > 0:
                time.sleep(min(timeout, 0.1))
            return False

        evt = getattr(self, "_preview_ready_event", None)
        if evt is None:
            deadline = None if timeout is None else (time.time() + max(0.0, float(timeout)))
            while True:
                if getattr(self, "_preview_last_token", 0) >= target:
                    return True
                if deadline is not None and time.time() >= deadline:
                    return False
                time.sleep(0.05)
        deadline = None if timeout is None else (time.time() + max(0.0, float(timeout)))
        while True:
            if getattr(self, "_preview_last_token", 0) >= target:
                return True
            if deadline is not None:
                remaining = deadline - time.time()
                if remaining <= 0:
                    return False
                wait_time = min(0.25, remaining)
            else:
                wait_time = 0.25
            try:
                evt.wait(wait_time)
            except Exception:
                return getattr(self, "_preview_last_token", 0) >= target
        return getattr(self, "_preview_last_token", 0) >= target

    def show_preview(self, img: Image.Image, wait_token=None):
        if img is None:
            self._signal_preview_ready(wait_token=wait_token)
            return
        try:
            w, h = img.size
            if w < 1 or h < 1:
                self.log_print(f"[Preview] skip display: empty image ({w}x{h})")
                return
            preview = img.copy()
            if preview.mode not in ("RGB", "RGBA"):
                preview = preview.convert("RGB")
            preview.thumbnail((720, 240))
            self.ocr_preview_imgtk = ImageTk.PhotoImage(preview)
            self.img_label.configure(image=self.ocr_preview_imgtk)
            try:
                self.img_label.update_idletasks()
            except Exception:
                pass
        except Exception as e:
            try:
                self.log_print(f"[Preview] display failed: {e}")
            except Exception:
                pass
        finally:
            self._signal_preview_ready(wait_token=wait_token)

    # ---------- Live preview (cursor tracking) ----------
    def toggle_live_preview(self):
        if self.live_preview_running:
            self.stop_live_preview()
        else:
            self.start_live_preview()

    def start_live_preview(self):
        if self.live_preview_running:
            return
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                self.log_print("Cannot start live preview without RDP connection.")
                return
        try:
            win = tk.Toplevel(self)
        except tk.TclError:
            self.log_print("Cannot open live preview window (no display available).")
            return

        win.title("Live Cursor Preview")
        win.geometry("640x360")
        win.attributes("-topmost", True)
        lbl = ttk.Label(win)
        lbl.pack(fill=tk.BOTH, expand=True)
        self.live_preview_window = win
        self.live_preview_label = lbl
        self.live_preview_running = True
        win.protocol("WM_DELETE_WINDOW", self.stop_live_preview)
        self._refresh_live_preview()
        self.log_print("Live preview started.")

    def stop_live_preview(self):
        was_running = self.live_preview_running
        self.live_preview_running = False
        if self.live_preview_window is not None:
            try:
                self.live_preview_window.destroy()
            except tk.TclError:
                pass
        self.live_preview_window = None
        self.live_preview_label = None
        self.live_preview_imgtk = None
        if was_running:
            self.log_print("Live preview stopped.")

    def _refresh_live_preview(self):
        if not self.live_preview_running:
            return
        if not self.live_preview_window or not self.live_preview_label:
            self.stop_live_preview()
            return
        rect = self.current_rect
        if not rect:
            self.stop_live_preview()
            return
        left, top, right, bottom = rect
        width, height = right - left, bottom - top
        try:
            cx, cy = pyautogui.position()
        except Exception:
            cx = cy = None

        view_size = 240
        cap_w = max(40, min(view_size, width))
        cap_h = max(40, min(view_size, height))
        if (
            cx is not None
            and cy is not None
            and left <= cx <= right
            and top <= cy <= bottom
        ):
            cap_left = int(min(max(cx - cap_w // 2, left), max(left, right - cap_w)))
            cap_top = int(min(max(cy - cap_h // 2, top), max(top, bottom - cap_h)))
            cursor_inside = True
        else:
            cap_left = int(left + max(0, (width - cap_w) // 2))
            cap_top = int(top + max(0, (height - cap_h) // 2))
            cursor_inside = False

        cap_w = int(min(cap_w, right - cap_left))
        cap_h = int(min(cap_h, bottom - cap_top))
        if cap_w <= 0 or cap_h <= 0:
            self.stop_live_preview()
            return

        try:
            img = grab_xywh(cap_left, cap_top, cap_w, cap_h)
        except Exception as exc:
            self.log_print(f"Live preview capture failed: {exc}")
            self.stop_live_preview()
            return

        draw = ImageDraw.Draw(img)
        draw.rectangle(
            [(0, 0), (img.width - 1, img.height - 1)], outline="yellow", width=2
        )

        info_text = "cursor outside"
        if cursor_inside and cx is not None and cy is not None:
            local_x = cx - cap_left
            local_y = cy - cap_top
            draw.line(
                [(local_x - 12, local_y), (local_x + 12, local_y)], fill="red", width=2
            )
            draw.line(
                [(local_x, local_y - 12), (local_x, local_y + 12)], fill="red", width=2
            )
            rel_x = (cx - left) / width if width else 0
            rel_y = (cy - top) / height if height else 0
            info_text = f"abs({cx},{cy}) rel({rel_x:.3f},{rel_y:.3f})"
        elif cx is not None and cy is not None:
            info_text = f"abs({cx},{cy})"

        draw.rectangle([(0, 0), (img.width, 18)], fill="black")
        draw.text((4, 2), info_text, fill="white")

        display = img.copy()
        display.thumbnail((360, 360))
        self.live_preview_imgtk = ImageTk.PhotoImage(display)
        if self.live_preview_label:
            self.live_preview_label.configure(image=self.live_preview_imgtk)
        if self.live_preview_window:
            self.live_preview_window.after(200, self._refresh_live_preview)

    def _reset_ocr_log_state(self):
        self._ocr_log_paths = {}

    def _append_ocr_log(self, log_label, section, lines, prefix="", keywords=None):
        if not log_label or not lines:
            return
        ensure_log_dir()
        if not hasattr(self, "_ocr_log_paths"):
            self._ocr_log_paths = {}
        if log_label not in self._ocr_log_paths:
            timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            safe = sanitize_filename(log_label)
            filename = f"{timestamp}_{safe}.log"
            path = os.path.join(LOG_DIR, filename)
            self._ocr_log_paths[log_label] = path
            info = f"{prefix}OCR log file for '{log_label}' → {path}"
            self.log_print(info.strip())
        path = self._ocr_log_paths.get(log_label)
        if not path:
            return
        try:
            with open(path, "a", encoding="utf-8") as fh:
                stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                fh.write(f"\n[{stamp}] Section: {section}\n")
                if keywords:
                    try:
                        joined = ", ".join(str(k) for k in keywords if str(k).strip())
                    except Exception:
                        joined = ""
                    if joined:
                        fh.write(f"Keywords: {joined}\n")
                for idx, entry in enumerate(lines, 1):
                    if isinstance(entry, (list, tuple)) and len(entry) == 5:
                        x, y, w, h, text = entry
                    else:
                        x = y = w = h = None
                        text = (
                            entry if not isinstance(entry, (list, tuple)) else entry[-1]
                        )
                    raw = text or ""
                    norm = normalize_line(raw)
                    soft = normalize_line_soft(raw)
                    fh.write(f"{idx:03d}: ({x},{y},{w},{h}) -> {raw}\n")
                    if norm and norm != raw:
                        fh.write(f"      norm: {norm}\n")
                    if soft and soft not in {raw, norm}:
                        fh.write(f"      soft: {soft}\n")
        except Exception as exc:
            self.log_print(f"{prefix}Failed to write OCR log for '{log_label}': {exc}")

    def log_print(self, text):
        self.log.insert(tk.END, str(text) + "\n")
        self.log.see(tk.END)
        self.update_idletasks()

    def clear_simple_log(self):
        if not hasattr(self, "simple_log"):
            return
        self.simple_log.configure(state="normal")
        self.simple_log.delete("1.0", tk.END)
        self.simple_log.configure(state="disabled")

    def simple_log_print(self, text):
        if not hasattr(self, "simple_log"):
            return
        self.simple_log.configure(state="normal")
        self.simple_log.insert(tk.END, str(text) + "\n")
        self.simple_log.see(tk.END)
        self.simple_log.configure(state="disabled")
        self.update_idletasks()

    def _render_simple_log_lines(self, lines):
        if not hasattr(self, "simple_log"):
            return
        self.simple_log.configure(state="normal")
        self.simple_log.delete("1.0", tk.END)
        for line in lines or []:
            if line is None:
                continue
            self.simple_log.insert(tk.END, str(line) + "\n")
        self.simple_log.configure(state="disabled")
        self.simple_log.see(tk.END)
        self.update_idletasks()

    def _should_skip_manual_waits(self):
        for attr in ("skip_waits_var", "rechnungen_skip_waits_var"):
            var = getattr(self, attr, None)
            if var is None:
                continue
            try:
                if bool(var.get()):
                    return True
            except Exception:
                continue
        return False

    def pull_form_into_cfg(self):
        self.cfg["rdp_title_regex"] = self.rdp_var.get().strip()
        self.cfg["excel_path"] = self.xls_var.get().strip()
        sv = self.sheet_var.get().strip()
        self.cfg["excel_sheet"] = int(sv) if sv.isdigit() else sv
        self.cfg["input_column"] = self.col_var.get().strip()
        self.cfg["results_csv"] = self.csv_var.get().strip()
        self.cfg["tesseract_path"] = self.tess_var.get().strip()
        self.cfg["tesseract_lang"] = self.lang_var.get().strip() or "deu+eng"
        try:
            self.cfg["type_delay"] = float(
                self.type_var.get() or DEFAULTS["type_delay"]
            )
        except:
            self.cfg["type_delay"] = DEFAULTS["type_delay"]
        try:
            self.cfg["post_search_wait"] = float(
                self.wait_var.get() or DEFAULTS["post_search_wait"]
            )
        except:
            self.cfg["post_search_wait"] = DEFAULTS["post_search_wait"]
        self.cfg["start_cell"] = (self.start_cell_var.get() or "").strip()
        try:
            self.cfg["max_rows"] = int(self.max_rows_var.get() or "0")
        except:
            self.cfg["max_rows"] = 0
        try:
            self.cfg["upscale_x"] = int(self.upscale_var.get() or "4")
        except:
            self.cfg["upscale_x"] = 4
        self.cfg["color_ocr"] = bool(self.color_var.get())
        self.cfg["use_full_region_parse"] = bool(self.fullparse_var.get())
        self.cfg["keyword"] = self.keyword_var.get().strip() or "Honorar"
        self.cfg["normalize_ocr"] = bool(self.normalize_var.get())

        # Profiles / selection
        self.cfg["use_amount_profile"] = bool(self.use_profile_var.get())
        self.cfg["active_amount_profile"] = self.profile_var.get() or ""

        # Fees settings
        if hasattr(self, "fees_search_var"):
            self.cfg["fees_search_token"] = self.fees_search_var.get().strip() or "KFB"
            self.cfg["fees_bad_prefixes"] = self.fees_bad_var.get().strip()
            try:
                self.cfg["fees_pages_max_clicks"] = int(
                    self.fees_pages_max_var.get() or "12"
                )
            except:
                self.cfg["fees_pages_max_clicks"] = 12
            self.cfg["fees_overlay_skip_waits"] = bool(self.fees_skip_waits_var.get())
            self.cfg["fees_csv_path"] = (
                self.fees_csv_var.get().strip() or "fees_results.csv"
            )

        self.cfg["includes"] = self.includes_var.get().strip()
        self.cfg["excludes"] = self.excludes_var.get().strip()
        self.cfg["exclude_prefix_k"] = bool(self.exclude_k_var.get())
        if hasattr(self, "ignore_top_doc_row_var"):
            self.cfg["ignore_top_doc_row"] = bool(self.ignore_top_doc_row_var.get())
        else:
            self.cfg["ignore_top_doc_row"] = False
        self.cfg["streitwert_term"] = self.streitwort_var.get().strip() or "Streitwert"
        try:
            self.cfg["doc_open_wait"] = float(self.docwait_var.get() or "1.2")
        except Exception:
            self.cfg["doc_open_wait"] = 1.2
        try:
            self.cfg["pdf_hit_wait"] = float(self.hitwait_var.get() or "1.0")
        except Exception:
            self.cfg["pdf_hit_wait"] = 1.0
        self.cfg["streitwert_results_csv"] = (
            self.streit_csv_var.get().strip() or "streitwert_results.csv"
        )
        if hasattr(self, "rechnungen_csv_var"):
            self.cfg["rechnungen_results_csv"] = (
                self.rechnungen_csv_var.get().strip()
                or "Streitwert_Results_Rechnungen.csv"
            )
        if hasattr(self, "rechnungen_only_csv_var"):
            self.cfg["rechnungen_only_results_csv"] = (
                self.rechnungen_only_csv_var.get().strip()
                or "rechnungen_only_results.csv"
            )
        if hasattr(self, "rechnungen_gg_csv_var"):
            self.cfg["rechnungen_gg_results_csv"] = (
                self.rechnungen_gg_csv_var.get().strip()
                or "rechnungen_gg_results.csv"
            )
        if hasattr(self, "rechnungen_search_wait_var"):
            try:
                self.cfg["rechnungen_search_wait"] = float(
                    self.rechnungen_search_wait_var.get() or "1.2"
                )
            except Exception:
                self.cfg["rechnungen_search_wait"] = DEFAULTS.get(
                    "rechnungen_search_wait", 1.2
                )
        if hasattr(self, "rechnungen_region_wait_var"):
            try:
                self.cfg["rechnungen_region_wait"] = float(
                    self.rechnungen_region_wait_var.get() or "0.0"
                )
            except Exception:
                self.cfg["rechnungen_region_wait"] = DEFAULTS.get(
                    "rechnungen_region_wait", 0.0
                )
        if hasattr(self, "rechnungen_skip_waits_var"):
            self.cfg["rechnungen_overlay_skip_waits"] = bool(
                self.rechnungen_skip_waits_var.get()
            )
        if hasattr(self, "log_dir_var"):
            log_dir = (self.log_dir_var.get() or "").strip()
            self.cfg["log_folder"] = log_dir or LOG_DIR
        if hasattr(self, "log_extract_csv_var"):
            log_csv = (self.log_extract_csv_var.get() or "").strip()
            default_csv = DEFAULTS.get(
                "log_extract_results_csv", "streitwert_log_extract.csv"
            )
            self.cfg["log_extract_results_csv"] = log_csv or default_csv
        self.cfg["streitwert_overlay_skip_waits"] = bool(self.skip_waits_var.get())
        if hasattr(self, "pdf_view_wait_var"):
            try:
                self.cfg["pdf_view_extra_wait"] = float(
                    self.pdf_view_wait_var.get() or "2.0"
                )
            except Exception:
                self.cfg["pdf_view_extra_wait"] = DEFAULTS["pdf_view_extra_wait"]

    def load_config(self):
        try:
            self.cfg = load_cfg()  # Load from file

            # Update form values from cfg
            self.rdp_var.set(self.cfg["rdp_title_regex"])
            self.xls_var.set(self.cfg["excel_path"])
            self.sheet_var.set(str(self.cfg["excel_sheet"]))
            self.col_var.set(self.cfg["input_column"])
            self.csv_var.set(self.cfg["results_csv"])
            self.tess_var.set(self.cfg["tesseract_path"])
            self.lang_var.set(self.cfg["tesseract_lang"])
            self.type_var.set(str(self.cfg["type_delay"]))
            self.wait_var.set(str(self.cfg["post_search_wait"]))
            self.start_cell_var.set(self.cfg.get("start_cell", ""))
            self.max_rows_var.set(str(self.cfg.get("max_rows", 0)))
            self.upscale_var.set(str(self.cfg.get("upscale_x", 4)))
            self.color_var.set(self.cfg.get("color_ocr", True))
            self.fullparse_var.set(self.cfg.get("use_full_region_parse", True))
            self.keyword_var.set(self.cfg.get("keyword", "Honorar"))
            self.normalize_var.set(self.cfg.get("normalize_ocr", True))
            self.includes_var.set(self.cfg.get("includes", "Urt,SWB,SW"))
            self.excludes_var.set(self.cfg.get("excludes", "SaM,KLE"))
            self.exclude_k_var.set(self.cfg.get("exclude_prefix_k", True))

            # Fees settings
            if hasattr(self, "fees_search_var"):
                self.fees_search_var.set(self.cfg.get("fees_search_token", "KFB"))
                self.fees_bad_var.set(
                    self.cfg.get("fees_bad_prefixes", "SVRAGS;SVR-AGS;Skrags;SV RAGS")
                )
                self.fees_pages_max_var.set(
                    str(self.cfg.get("fees_pages_max_clicks", 12))
                )
                self.fees_skip_waits_var.set(
                    self.cfg.get("fees_overlay_skip_waits", True)
                )
                self.fees_csv_var.set(self.cfg.get("fees_csv_path", "fees_results.csv"))

            if hasattr(self, "ignore_top_doc_row_var"):
                self.ignore_top_doc_row_var.set(
                    self.cfg.get("ignore_top_doc_row", False)
                )
            self.streitwort_var.set(self.cfg.get("streitwert_term", "Streitwert"))
            self.docwait_var.set(str(self.cfg.get("doc_open_wait", 1.2)))
            self.hitwait_var.set(str(self.cfg.get("pdf_hit_wait", 1.0)))
            self.streit_csv_var.set(
                self.cfg.get("streitwert_results_csv", "streitwert_results.csv")
            )
            if hasattr(self, "rechnungen_csv_var"):
                self.rechnungen_csv_var.set(
                    self.cfg.get(
                        "rechnungen_results_csv",
                        "Streitwert_Results_Rechnungen.csv",
                    )
                )
            if hasattr(self, "rechnungen_only_csv_var"):
                self.rechnungen_only_csv_var.set(
                    self.cfg.get(
                        "rechnungen_only_results_csv",
                        "rechnungen_only_results.csv",
                    )
                )
            if hasattr(self, "rechnungen_gg_csv_var"):
                self.rechnungen_gg_csv_var.set(
                    self.cfg.get(
                        "rechnungen_gg_results_csv",
                        "rechnungen_gg_results.csv",
                    )
                )
            if hasattr(self, "rechnungen_search_wait_var"):
                wait_val = self.cfg.get(
                    "rechnungen_search_wait",
                    self.cfg.get("post_search_wait", 1.2),
                )
                self.rechnungen_search_wait_var.set(str(wait_val))
            if hasattr(self, "rechnungen_region_wait_var"):
                region_wait = self.cfg.get(
                    "rechnungen_region_wait",
                    DEFAULTS.get("rechnungen_region_wait", 0.0),
                )
                self.rechnungen_region_wait_var.set(str(region_wait))
            if hasattr(self, "rechnungen_skip_waits_var"):
                self.rechnungen_skip_waits_var.set(
                    self.cfg.get("rechnungen_overlay_skip_waits", False)
                )
            if hasattr(self, "log_dir_var"):
                self.log_dir_var.set(self.cfg.get("log_folder", LOG_DIR))
            if hasattr(self, "log_extract_csv_var"):
                self.log_extract_csv_var.set(
                    self.cfg.get(
                        "log_extract_results_csv", "streitwert_log_extract.csv"
                    )
                )
            if hasattr(self, "skip_waits_var"):
                self.skip_waits_var.set(
                    self.cfg.get("streitwert_overlay_skip_waits", False)
                )
            if hasattr(self, "pdf_view_wait_var"):
                self.pdf_view_wait_var.set(
                    str(self.cfg.get("pdf_view_extra_wait", 2.0))
                )
            hits_pt = self.cfg.get("pdf_hits_point")
            if not (isinstance(hits_pt, (list, tuple)) and len(hits_pt) == 2):
                legacy = self.cfg.get("pdf_hits_region")
                if (
                    isinstance(legacy, (list, tuple))
                    and len(legacy) == 4
                    and all(isinstance(v, (int, float)) for v in legacy)
                ):
                    converted = [legacy[0] + legacy[2] / 2, legacy[1] + legacy[3] / 2]
                    self.cfg["pdf_hits_point"] = converted
                    hits_pt = converted
                    self.cfg.pop("pdf_hits_region", None)
            if hasattr(self, "pdf_hits_var"):
                if (
                    isinstance(hits_pt, (list, tuple))
                    and len(hits_pt) == 2
                    and all(isinstance(v, (int, float)) for v in hits_pt)
                ):
                    self.pdf_hits_var.set(f"{hits_pt[0]:.3f}, {hits_pt[1]:.3f}")
                else:
                    self.pdf_hits_var.set("")
            hits2_pt = self.cfg.get("pdf_hits_second_point")
            if hasattr(self, "pdf_hits2_var"):
                if (
                    isinstance(hits2_pt, (list, tuple))
                    and len(hits2_pt) == 2
                    and all(isinstance(v, (int, float)) for v in hits2_pt)
                ):
                    self.pdf_hits2_var.set(f"{hits2_pt[0]:.3f}, {hits2_pt[1]:.3f}")
                else:
                    self.pdf_hits2_var.set("")
            hits3_pt = self.cfg.get("pdf_hits_third_point")
            if hasattr(self, "pdf_hits3_var"):
                if (
                    isinstance(hits3_pt, (list, tuple))
                    and len(hits3_pt) == 2
                    and all(isinstance(v, (int, float)) for v in hits3_pt)
                ):
                    self.pdf_hits3_var.set(f"{hits3_pt[0]:.3f}, {hits3_pt[1]:.3f}")
                else:
                    self.pdf_hits3_var.set("")
            view_pt = self.cfg.get("doc_view_point")
            if isinstance(view_pt, (list, tuple)) and len(view_pt) == 2:
                self.doc_view_var.set(f"{view_pt[0]:.3f}, {view_pt[1]:.3f}")
            else:
                self.doc_view_var.set("")
            close_pt = self.cfg.get("pdf_close_point")
            if hasattr(self, "pdf_close_var"):
                if (
                    isinstance(close_pt, (list, tuple))
                    and len(close_pt) == 2
                    and all(isinstance(v, (int, float)) for v in close_pt)
                ):
                    self.pdf_close_var.set(f"{close_pt[0]:.3f}, {close_pt[1]:.3f}")
                else:
                    self.pdf_close_var.set("")

            if hasattr(self, "rechnungen_region_var"):
                rechn_box = self.cfg.get("rechnungen_region")
                if (
                    isinstance(rechn_box, (list, tuple))
                    and len(rechn_box) == 4
                    and all(isinstance(v, (int, float)) for v in rechn_box)
                ):
                    self.rechnungen_region_var.set(
                        f"{rechn_box[0]:.3f}, {rechn_box[1]:.3f}, {rechn_box[2]:.3f}, {rechn_box[3]:.3f}"
                    )
                else:
                    self.rechnungen_region_var.set("")
            if hasattr(self, "rechnungen_gg_region_var"):
                gg_box = self.cfg.get("rechnungen_gg_region")
                if (
                    isinstance(gg_box, (list, tuple))
                    and len(gg_box) == 4
                    and all(isinstance(v, (int, float)) for v in gg_box)
                ):
                    self.rechnungen_gg_region_var.set(
                        f"{gg_box[0]:.3f}, {gg_box[1]:.3f}, {gg_box[2]:.3f}, {gg_box[3]:.3f}"
                    )
                else:
                    self.rechnungen_gg_region_var.set("")

            # Profiles UI
            self.profile_names = [
                p["name"] for p in self.cfg.get("amount_profiles", [])
            ]
            self.profile_box["values"] = self.profile_names
            self.profile_var.set(self.cfg.get("active_amount_profile", ""))
            self.use_profile_var.set(self.cfg.get("use_amount_profile", False))
            self._refresh_profile_fields_from_active()

            self.log_print("Configuration loaded successfully")
            messagebox.showinfo("Load Config", "Configuration loaded successfully")
        except Exception as e:
            self.log_print(f"Error loading configuration: {e}")
            messagebox.showerror("Load Config", f"Error loading configuration: {e}")

    # ----------- Profiles helpers -----------
    def _get_active_profile(self):
        name = self.profile_var.get().strip()
        for p in self.cfg.get("amount_profiles", []):
            if p.get("name") == name:
                return p
        return None

    def _refresh_profile_fields_from_active(self):
        prof = self._get_active_profile()
        if prof:
            self.new_prof_name_var.set(prof.get("name", ""))
            self.prof_keyword_var.set(prof.get("keyword", "") or "")
            self._current_profile_sub_region = prof.get("sub_region", None)
        else:
            self.prof_keyword_var.set("")
            self._current_profile_sub_region = None

    # --- Fees helpers ---

    def _fees_should_skip(self, line: str) -> bool:
        """Skip if line starts with any bad prefix (case-insensitive)."""
        bad = (self.fees_bad_var.get() or self.cfg.get("fees_bad_prefixes", "")).strip()
        if not bad:
            return False
        toks = [b.strip().lower() for b in bad.split(";") if b.strip()]
        if not toks:
            return False
        line_raw = (line or "").strip()
        line_lower = line_raw.lower()
        line_norm = normalize_line_soft(line_raw).lower()

        for tok in toks:
            if not tok:
                continue
            tok_norm = normalize_line_soft(tok).lower()
            if line_lower.startswith(tok) or line_norm.startswith(tok_norm):
                return True
        return False

    def _click_file_search_and_type_kfb(self):
        """Click into fees_file_search_region and type the token (once)."""
        try:
            ax, ay, aw, ah = rel_to_abs(
                self.current_rect, self.cfg.get("fees_file_search_region", [0, 0, 0, 0])
            )
        except Exception:
            self.log_print("[Fees] File search region not configured.")
            return
        x = int(ax + max(1, aw // 2))
        y = int(ay + max(1, ah // 2))
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass
        pyautogui.click(x, y)
        time.sleep(0.1)
        token = (
            self.fees_search_var.get() or self.cfg.get("fees_search_token", "KFB")
        ).strip()
        self._type_fast(token)
        pyautogui.press("enter")

    def _type_fast(self, s):
        delay = max(
            0.002,
            float(
                self.typing_delay_var.get()
                if hasattr(self, "typing_delay_var")
                else 0.008
            ),
        )
        pyautogui.typewrite(s, interval=delay)

    def _fees_overlay_wait(self, what):
        """Overlay-preferring waits like Streitwert."""
        if self.fees_skip_waits_var.get() or self.cfg.get(
            "fees_overlay_skip_waits", True
        ):
            if what == "doclist":
                self._wait_for_doclist_ready(prefix="[Fees] ")
            elif what == "pdf":
                self._wait_for_pdf_ready(prefix="[Fees] ")
        else:
            time.sleep(0.6)  # tiny safety fallback

    def _fees_analyze_seiten_region(
        self, x0, y0, W, H, max_clicks=None, img=None, lang=None
    ):
        """
        Inspect the configured Seiten strip and return:
          - preview image (for reuse by caller)
          - compact OCR summary text (bottom digits strip)
          - detected page-label sequence (e.g. "1 2 3 4")
          - click positions derived from the page labels
        """

        result = {
            "img": None,
            "token_summary": "",
            "digit_summary": "",
            "positions": [],
        }

        if W <= 0 or H <= 0:
            return result

        if lang is None:
            try:
                lang = self.lang_var.get().strip() or "deu+eng"
            except Exception:
                lang = "deu+eng"

        try:
            preview = img or self._grab_region_color(
                x0, y0, W, H, upscale_x=self.upscale_var.get()
            )
        except Exception:
            return result

        result["img"] = preview

        if preview.width < 2 or preview.height < 2:
            return result

        band_top = int(preview.height * 0.55)
        band = self._safe_crop(preview, (0, band_top, preview.width, preview.height))
        if band.width < 2 or band.height < 2:
            return result

        scale = 3
        proc = band.resize((band.width * scale, band.height * scale), Image.LANCZOS)
        proc = ImageOps.autocontrast(ImageOps.grayscale(proc))

        try:
            df = do_ocr_data(proc, lang=lang, psm=6)
        except Exception:
            df = None

        if df is None or "text" not in df.columns:
            return result

        texts = []
        digits = []
        for row in df.itertuples():
            text = str(getattr(row, "text", "")).strip()
            if not text:
                continue

            left = getattr(row, "left", 0)
            top = getattr(row, "top", 0)
            width = getattr(row, "width", 0)
            height = getattr(row, "height", 0)

            if pd.isna(left):
                left = 0
            if pd.isna(top):
                top = 0
            if pd.isna(width) or width <= 0:
                width = 1
            if pd.isna(height) or height <= 0:
                height = 1

            # Project coordinates back to the preview (pre-scale, pre-crop)
            left = float(left) / scale
            top = float(top) / scale + band_top
            width = float(width) / scale
            height = float(height) / scale

            texts.append(text)

            if re.fullmatch(r"\d+", text):
                center_x = left + width / 2.0
                center_y = top + height / 2.0
                digits.append((text, center_x, center_y, height))

        if texts:
            summary = " | ".join(texts)
            if len(summary) > 200:
                summary = summary[:197] + "..."
            result["token_summary"] = summary

        if not digits:
            return result

        digits.sort(key=lambda item: item[1])
        page_labels = []
        positions = []
        limit = max_clicks or len(digits)

        for idx, (label, cx, cy, h) in enumerate(digits):
            if idx >= limit:
                break
            abs_x = int(round(x0 + cx))
            baseline = y0 + cy
            target_y = baseline - h * 2.0
            target_y = max(y0 + 5, min(y0 + H - 5, target_y))
            positions.append((idx + 1, abs_x, int(round(target_y))))
            page_labels.append(label)

        result["digit_summary"] = " ".join(page_labels)
        result["positions"] = positions
        return result

    def _fees_iter_click_pages(self, max_clicks=None, return_positions=False):
        """Click across the Seiten thumbnails from left to right."""
        max_clicks = (
            int(self.cfg.get("fees_pages_max_clicks", 12))
            if max_clicks is None
            else max_clicks
        )
        try:
            x0, y0, W, H = rel_to_abs(
                self.current_rect, self.cfg.get("fees_seiten_region", [0, 0, 0, 0])
            )
        except Exception:
            self.log_print("[Fees] Seiten region not configured.")
            return
        if W <= 0 or H <= 0:
            self.log_print("[Fees] Seiten region not configured.")
            return

        analysis = self._fees_analyze_seiten_region(x0, y0, W, H, max_clicks=max_clicks)
        positions = analysis.get("positions") or []

        if not positions:
            # Fallback to evenly spaced clicks if OCR failed to find page labels
            step = max(1, W // max(1, max_clicks))
            positions = [
                (i + 1, x0 + step // 2 + i * step, y0 + H // 2)
                for i in range(max_clicks)
            ]

        for idx, x, y in positions:
            pyautogui.click(x, y)
            time.sleep(0.15)
            self._fees_overlay_wait("pdf")
            if return_positions:
                yield (idx, x, y)
            else:
                yield idx

    def _is_pdf_open(self):
        """
        Lightweight check to see if a PDF view is open in DATEV.
        Uses the calibrated pdf_text_region and looks for visible text.
        Returns True if some text (non-empty OCR) is present, else False.
        """
        if not self._has("pdf_text_region"):
            return False

        try:
            # Convert relative pdf_text_region to absolute coords
            x, y, w, h = rel_to_abs(self.current_rect, self._get("pdf_text_region"))
            if w <= 0 or h <= 0:
                return False
            img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
            df = do_ocr_data(
                img, lang=(self.lang_var.get().strip() or "deu+eng"), psm=6
            )
            if "text" not in df.columns:
                return False
            # any visible non-empty token = PDF is open
            texts = [t for t in df["text"].tolist() if str(t).strip()]
            return len(texts) > 3
        except Exception:
            return False

    def _fees_scan_current_page_amount(self):
        """OCR pdf_text_region and return (numeric_amount_display) if page has both numeric and words."""
        if not self._has("pdf_text_region"):
            self.log_print("[Fees] pdf_text_region not set; calibrate in Streitwert.")
            return None
        # Convert relative pdf_text_region to absolute coords
        x, y, w, h = rel_to_abs(self.current_rect, self._get("pdf_text_region"))
        if w <= 0 or h <= 0:
            return None
        img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
        df = do_ocr_data(img, lang=(self.lang_var.get().strip() or "deu+eng"), psm=6)
        try:
            text = " ".join([t for t in df["text"].tolist() if str(t).strip()]).strip()
        except Exception:
            text = ""
        if not text:
            return None

        has_words = bool(self._WORDS_HINT_RE.search(text))
        m = self._AMT_NUM_RE.search(text)
        if m and has_words:
            amt = m.group(0)
            return amt
        return None

    def _fees_open_and_extract_one(self, row_idx, prefix=""):
        """Open row_idx (0-based) that looks like a KFB, scan pages for amount, close PDF, return string or None."""
        # 1) Click row, hit View (use your Streitwert 'View' button config)
        if not self._click_doclist_row(row_idx):  # use your existing helper
            self.log_print(f"{prefix}Cannot click row {row_idx}.")
            return None
        if not self._click_view_button(prefix=prefix):
            return None
        self._fees_overlay_wait("pdf")

        # 2) Click through pages to find amount
        amount = None
        for _ in self._fees_iter_click_pages():
            val = self._fees_scan_current_page_amount()
            if val:
                amount = val
                break

        # 3) Close PDF (reuse Streitwert close)
        self._close_active_pdf(prefix=prefix)
        self._fees_overlay_wait("doclist")
        return amount

    def _fees_is_kfb_line(self, text: str) -> bool:
        if not text:
            return False
        if self._KFB_RE.search(text):
            return True
        norm = normalize_line_soft(text).lower()
        if not norm:
            return False
        if self._KFB_WORD_RE.search(norm):
            return True
        compact = re.sub(r"[^a-zß]", "", norm)
        return compact.startswith("kostenfestsetzungsbeschl")

    def _fees_collect_kfb_rows(self):
        """Return list of (row_index, row_text) that look like KFB and not skipped by bad prefixes."""
        rows_with_boxes = self._ocr_doclist_rows_boxes()
        kfb = []
        for i, (line, _) in enumerate(rows_with_boxes):
            s = (line or "").strip()
            if not s:
                continue
            if self._fees_should_skip(s):
                continue
            if self._fees_is_kfb_line(s):
                kfb.append((i, s))
        return kfb

    def edit_fees_bad_prefixes(self):
        """Prompt the user to edit the semicolon-separated bad prefixes list."""
        try:
            current = (self.fees_bad_var.get() or "").strip()
        except Exception:
            current = ""
        value = simpledialog.askstring(
            "Fees Bad Prefixes",
            "Prefixes to skip (separate with semicolons):",
            initialvalue=current,
            parent=self,
        )
        if value is not None:
            self.fees_bad_var.set(value.strip())

    def pick_fees_file_search_region(self):
        """Two-click calibration for the KFB search text region."""
        rb = self._two_click_box(
            "Hover TOP-LEFT of the file search box area, then press OK.",
            "Hover BOTTOM-RIGHT of the file search box area, then press OK.",
        )
        if rb:
            self.cfg["fees_file_search_region"] = rb
            self.fees_file_search_var.set(
                f"{rb[0]:.3f}, {rb[1]:.3f}, {rb[2]:.3f}, {rb[3]:.3f}"
            )
            self.log_print(f"[Fees] File search region set: {rb}")

    def pick_fees_seiten_region(self):
        """Two-click calibration for the thumbnails strip."""
        rb = self._two_click_box(
            "Hover TOP-LEFT of the Seiten/Pages thumbnails strip, then press OK.",
            "Hover BOTTOM-RIGHT of the Seiten/Pages thumbnails strip, then press OK.",
        )
        if rb:
            self.cfg["fees_seiten_region"] = rb
            self.fees_seiten_var.set(
                f"{rb[0]:.3f}, {rb[1]:.3f}, {rb[2]:.3f}, {rb[3]:.3f}"
            )
            self.log_print(f"[Fees] Seiten region set: {rb}")

    def run_fees(self):
        prefix = "[Fees] "
        try:
            self.apply_paths_to_tesseract()
        except Exception:
            pass

        # Instance detection first (ensures we know how many KFBs to open)
        inst_info = self.detect_instance(prefix=prefix) or {}
        inst = inst_info.get("instance") or 1
        self.log_print(f"{prefix}Instance → open first {inst} KFB file(s).")

        # Put KFB into file-search (once)
        self._click_file_search_and_type_kfb()
        self._fees_overlay_wait("doclist")

        # Find KFB rows in doclist
        kfb_rows = self._fees_collect_kfb_rows()
        if not kfb_rows:
            self.log_print(f"{prefix}No KFB entries found.")
            return
        self.log_print(f"{prefix}Found {len(kfb_rows)} KFB entr{'y' if len(kfb_rows)==1 else 'ies'}.")

        # Open first N, extract
        N = min(inst, len(kfb_rows))
        amounts = [None, None, None]  # inst1, inst2, inst3
        for j in range(N):
            row_idx, line = kfb_rows[j]
            self.log_print(f"{prefix}Opening KFB {j+1}/{N}: row {row_idx} → {line}")
            amt = self._fees_open_and_extract_one(row_idx, prefix=prefix)
            if amt:
                self.log_print(f"{prefix}Amount: {amt}")
            else:
                self.log_print(f"{prefix}Amount not found.")
            amounts[j] = amt

        # Build CSV row
        aktenzeichen = (
            self._current_aktenzeichen_text()
            if hasattr(self, "_current_aktenzeichen_text")
            else ""
        )
        row = {
            "aktenzeichen": aktenzeichen,
            "instance_detected": inst,
            "fees_inst1": amounts[0] or "",
            "fees_inst2": amounts[1] or "",
            "fees_inst3": amounts[2] or "",
        }

        # Write CSV
        path = self.cfg.get("fees_csv_path", "fees_results.csv")
        write_header = not os.path.exists(path)
        with open(path, "a", encoding="utf-8", newline="") as f:
            import csv

            w = csv.DictWriter(f, fieldnames=list(row.keys()))
            if write_header:
                w.writeheader()
            w.writerow(row)
        self.log_print(f"{prefix}Saved → {path}")

    def test_fees(self):
        prefix = "[Fees Test] "
        try:
            self.apply_paths_to_tesseract()
        except Exception:
            pass

        self.detect_instance(prefix=prefix)

        amount = None
        pdf_open = False
        try:
            pdf_open = self._is_pdf_open()
        except Exception:
            pdf_open = False

        if pdf_open:
            for _ in self._fees_iter_click_pages(max_clicks=6):
                v = self._fees_scan_current_page_amount()
                if v:
                    amount = v
                    break
        else:
            self._click_file_search_and_type_kfb()
            self._fees_overlay_wait("doclist")
            kfb_rows = self._fees_collect_kfb_rows()
            if kfb_rows:
                self.log_print(
                    f"{prefix}Found {len(kfb_rows)} KFB entr{'y' if len(kfb_rows)==1 else 'ies'} for test."
                )
                amount = self._fees_open_and_extract_one(kfb_rows[0][0], prefix=prefix)

        self.log_print(f"{prefix}Amount on some page: {amount or '(none found)'}")

    def test_fees_seiten_clicks(self):
        prefix = "[Fees Seiten Test] "
        try:
            self.apply_paths_to_tesseract()
        except Exception:
            pass

        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                self.log_print(f"{prefix}No RDP connection.")
                return

        if not self._has("fees_seiten_region"):
            self.log_print(f"{prefix}Seiten region not configured.")
            return

        try:
            x, y, w, h = rel_to_abs(self.current_rect, self._get("fees_seiten_region"))
        except Exception:
            self.log_print(f"{prefix}Seiten region not configured.")
            return

        img = None
        try:
            img = self._grab_region_color(x, y, w, h, upscale_x=self.upscale_var.get())
            self.show_preview(img)
        except Exception:
            self.log_print(f"{prefix}Preview unavailable.")

        lang = None
        try:
            lang = self.lang_var.get().strip() or "deu+eng"
        except Exception:
            lang = "deu+eng"

        analysis = self._fees_analyze_seiten_region(
            x, y, w, h, max_clicks=6, img=img, lang=lang
        )

        summary = analysis.get("token_summary") or ""
        self.log_print(f"{prefix}OCR text: {summary or '(none)'}")

        digits_summary = analysis.get("digit_summary") or ""
        if digits_summary:
            self.log_print(f"{prefix}Detected pages: {digits_summary}")
        else:
            self.log_print(f"{prefix}Detected pages: (none)")

        positions = analysis.get("positions") or []
        clicks = 0
        if positions:
            for idx, cx, cy in positions:
                clicks = idx
                pyautogui.click(cx, cy)
                time.sleep(0.15)
                self._fees_overlay_wait("pdf")
                self.log_print(f"{prefix}Click {idx} at ({cx}, {cy}).")
        else:
            iterator = self._fees_iter_click_pages(max_clicks=6, return_positions=True)
            if iterator is None:
                return
            for idx, cx, cy in iterator:
                clicks = idx
                self.log_print(f"{prefix}Click {idx} at ({cx}, {cy}).")

        if clicks == 0:
            self.log_print(f"{prefix}No clicks executed.")

    def save_amount_profile(self):
        name = (self.new_prof_name_var.get() or "").strip()
        if not name:
            messagebox.showwarning("Save Profile", "Please enter a profile name.")
            return
        kw = (self.prof_keyword_var.get() or "").strip()
        sub = self._current_profile_sub_region
        # Update or add
        updated = False
        for p in self.cfg.setdefault("amount_profiles", []):
            if p.get("name") == name:
                p["keyword"] = kw
                if sub is not None:
                    p["sub_region"] = sub
                updated = True
                break
        if not updated:
            self.cfg["amount_profiles"].append(
                {"name": name, "keyword": kw, "sub_region": sub}
            )

        self.cfg["active_amount_profile"] = name
        self.profile_names = [p["name"] for p in self.cfg["amount_profiles"]]
        self.profile_box["values"] = self.profile_names
        self.profile_var.set(name)
        save_cfg(self.cfg)
        self.log_print(f"Profile '{name}' saved.")
        messagebox.showinfo("Profile", f"Profile '{name}' saved.")

    def delete_amount_profile(self):
        name = (self.profile_var.get() or "").strip()
        if not name:
            messagebox.showwarning("Delete Profile", "No active profile selected.")
            return
        profs = self.cfg.get("amount_profiles", [])
        new_list = [p for p in profs if p.get("name") != name]
        if len(new_list) == len(profs):
            messagebox.showinfo("Delete Profile", "Profile not found.")
            return
        self.cfg["amount_profiles"] = new_list
        if self.cfg.get("active_amount_profile") == name:
            self.cfg["active_amount_profile"] = ""
            self.profile_var.set("")
            self.new_prof_name_var.set("")
            self.prof_keyword_var.set("")
            self._current_profile_sub_region = None
        self.profile_names = [p["name"] for p in self.cfg["amount_profiles"]]
        self.profile_box["values"] = self.profile_names
        save_cfg(self.cfg)
        self.log_print(f"Profile '{name}' deleted.")
        messagebox.showinfo("Profile", f"Profile '{name}' deleted.")

    def save_config(self):
        try:
            self.pull_form_into_cfg()  # Update cfg from form values
            save_cfg(self.cfg)  # Save to file
            self.log_print("Configuration saved successfully")
            messagebox.showinfo("Save Config", "Configuration saved successfully")
        except Exception as e:
            self.log_print(f"Error saving configuration: {e}")
            messagebox.showerror("Save Config", f"Error saving configuration: {e}")


if __name__ == "__main__":
    app = RDPApp()
    app.mainloop()
