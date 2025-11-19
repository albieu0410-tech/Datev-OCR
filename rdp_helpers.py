import time
import threading
import re
import unicodedata
from decimal import Decimal, InvalidOperation

import numpy as np
import pyautogui
from PIL import Image
from mss import mss

try:  # Pywinauto renamed Desktop helper across releases
    from pywinauto.desktop import Desktop
except ImportError:
    try:
        from pywinauto.application import Desktop
    except ImportError:
        from pywinauto import Desktop
import pytesseract

try:
    import cv2
    _HAS_CV2 = True
except Exception:  # pragma: no cover - optional dependency
    cv2 = None
    _HAS_CV2 = False

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


GG_LABEL_TRANSLATE = str.maketrans(
    {"6": "G", "0": "G", "O": "G", "Q": "G", "C": "G", "€": "G"}
)
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
