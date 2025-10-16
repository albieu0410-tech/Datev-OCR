import os, io, json, time, threading, re, unicodedata
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageStat, ImageDraw
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
DEFAULTS = {
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
    # --- Rechnungen workflow (NEW) ---
    "rechnungen_region": [0.55, 0.30, 0.35, 0.40],
    "rechnungen_results_csv": "Streitwert_Results_Rechnungen.csv",
    "rechnungen_only_results_csv": "rechnungen_only_results.csv",
}
CFG_FILE = "rdp_automation_config.json"


# ------------------ Helpers ------------------
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


def grab_xywh(x, y, w, h):
    with mss() as sct:
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


# ---------- OCR TSV helpers (Streitwert) ----------
def normalize_line(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\u0080", "€")
    fix = str.maketrans({"O": "0", "o": "0", "S": "5", "s": "5", "l": "1", "I": "1", "B": "8"})
    t = text.translate(fix)
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\beur\b", "EUR", t, flags=re.IGNORECASE)
    return t


AMOUNT_RE = re.compile(
    r"(?:€\s*)?(?:\d{1,3}(?:[.\s]\d{3})+|\d+),\d{2}(?:\s*(?:EUR|€))?",
    re.IGNORECASE,
)
DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
INVOICE_RE = re.compile(r"\b\d{6,}\b")


def extract_amount_from_text(text: str):
    t = normalize_line(text)
    matches = list(AMOUNT_RE.finditer(t))
    if not matches:
        return None
    for match in matches:
        candidate = match.group(0).strip().strip(".,;: ")
        if re.search(r"(EUR|€)", candidate, re.IGNORECASE):
            return candidate
    return matches[0].group(0).strip().strip(".,;: ")


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
        remainder = amt[leading.end():].strip()
        remainder = re.sub(r"(\d)(EUR|€)$", r"\1 \2", remainder, flags=re.IGNORECASE)
        remainder = re.sub(r"\s+(EUR|€)$", r" \1", remainder, flags=re.IGNORECASE)
        if remainder and AMOUNT_RE.fullmatch(remainder):
            amt = remainder
    return amt


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
DIGIT_FIX = str.maketrans(
    {"O": "0", "o": "0", "S": "5", "s": "5", "l": "1", "I": "1", "B": "8"}
)


def normalize_token_digits(tok: str) -> str:
    return tok.translate(DIGIT_FIX)


def normalize_line_soft(text: str) -> str:
    if not text:
        return text
    text = (
        text.replace("0", "o")
        .replace("O", "o")
        .replace("1", "l")
        .replace("5", "s")
        .replace("B", "8")
    )
    parts = re.split(r"(\s+)", text)
    parts = [
        normalize_token_digits(p) if i % 2 == 0 else p for i, p in enumerate(parts)
    ]
    t = "".join(parts)
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\beur\b", "EUR", t, flags=re.IGNORECASE)
    t = re.sub(r"(\d+)\.(\d{2})\b", r"\1,\2", t)
    return t
def extract_amount_from_lines(lines, keyword=None):
    if not lines:
        return None, None

    processed = []
    for entry in lines:
        if isinstance(entry, (list, tuple)) and len(entry) == 5:
            _, y, _, _, text = entry
        else:
            y, text = entry
        processed.append((y, text))

    norm_lines = [(y, normalize_line_soft(t)) for y, t in processed]

    def find_amounts(text):
        amts = AMOUNT_RE.findall(text)
        non_zero = [
            a for a in amts if not a.startswith("0,") and not a.startswith("0.")
        ]
        return non_zero if non_zero else amts

    if keyword:
        k = normalize_line_soft(keyword.strip()).lower()
        for idx, (_, t) in enumerate(norm_lines):
            if k in t.lower():
                amts = find_amounts(t)
                if amts:
                    return amts[-1], t

    if keyword:
        k = normalize_line_soft(keyword.strip()).lower()
        for idx, (_, t) in enumerate(norm_lines):
            if k in t.lower():
                for offset in [0, 1, -1]:
                    check_idx = idx + offset
                    if 0 <= check_idx < len(norm_lines):
                        amts = find_amounts(norm_lines[check_idx][1])
                        if amts:
                            return amts[-1], norm_lines[check_idx][1]

    for _, t in norm_lines:
        amts = find_amounts(t)
        if amts:
            return amts[-1], t

    return None, None


# ------------------ App Class ------------------
class RDPApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RDP Automation (Tkinter)")
        self.geometry("1220x860")
        self.minsize(1080, 760)

        self.cfg = load_cfg()
        self.current_rect = None
        self.ocr_preview_imgtk = None
        self._current_profile_sub_region = None
        self.capture_countdown_seconds = 3
        self.live_preview_window = None
        self.live_preview_label = None
        self.live_preview_imgtk = None
        self.live_preview_running = False

        self.create_widgets()

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
        ocr_tab = ttk.Frame(notebook)

        notebook.add(general_tab, text="General")
        notebook.add(calibration_tab, text="Calibration")
        notebook.add(streit_tab, text="Streitwert")
        notebook.add(rechn_tab, text="Rechnungen")
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

        ttk.Label(excel_frame, text="(Optional) Input column name (if Start cell empty)").pack(
            anchor="w", pady=(6, 0)
        )
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
        self.includes_var = tk.StringVar(
            value=self.cfg.get("includes", "Urt,SWB,SW")
        )
        ttk.Entry(streit_frame, textvariable=self.includes_var, width=40).pack(
            anchor="w"
        )

        ttk.Label(streit_frame, text="Exclude tokens (comma-separated)").pack(
            anchor="w", pady=(6, 0)
        )
        self.excludes_var = tk.StringVar(
            value=self.cfg.get("excludes", "SaM,KLE")
        )
        ttk.Entry(streit_frame, textvariable=self.excludes_var, width=40).pack(
            anchor="w"
        )

        self.exclude_k_var = tk.BooleanVar(
            value=self.cfg.get("exclude_prefix_k", True)
        )
        ttk.Checkbutton(
            streit_frame,
            text="Exclude rows starting with 'K'",
            variable=self.exclude_k_var,
        ).pack(anchor="w", pady=(6, 0))

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
        self.docwait_var = tk.StringVar(
            value=str(self.cfg.get("doc_open_wait", 1.2))
        )
        ttk.Entry(row3, textvariable=self.docwait_var, width=6).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Label(row3, text="Hit wait (s)").pack(side=tk.LEFT)
        self.hitwait_var = tk.StringVar(
            value=str(self.cfg.get("pdf_hit_wait", 1.0))
        )
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
            value=self.cfg.get(
                "streitwert_results_csv", "streitwert_results.csv"
            )
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
        ttk.Entry(
            streit_frame, textvariable=self.rechnungen_csv_var, width=40
        ).pack(anchor="w", pady=(0, 4))

        ttk.Button(
            streit_frame,
            text="Test Streitwert Setup",
            command=self.test_streitwert_setup,
        ).pack(anchor="w", pady=(6, 0))

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
        rechn_frame = ttk.LabelFrame(
            rechn_tab, text="Rechnungen Calibration & Test"
        )
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
            text="Test Rechnungen Extraction",
            command=self.test_rechnungen_threaded,
        ).pack(anchor="w", pady=(6, 0))

        ttk.Label(rechn_frame, text="Rechnungen-only CSV").pack(
            anchor="w", pady=(6, 0)
        )
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
            "Pick Result Region", "Position your mouse over the TOP-LEFT corner of the region."
        )

        x2, y2 = self._prompt_and_capture_point(
            "Pick Result Region", "Now position your mouse over the BOTTOM-RIGHT corner."
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
        return abs_to_rel(
            self.current_rect, abs_box=(left, top, width, height)
        )

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

    def _grab_region_color(self):
        self._validate_result_region()
        rx, ry, rw, rh = rel_to_abs(self.current_rect, self.cfg["result_region"])
        region = grab_xywh(rx, ry, rw, rh)
        scale = max(1, int(self.upscale_var.get() or 3))
        return upscale_pil(region, scale=scale)

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

            summary = self._extract_rechnungen_summary(prefix="[Rechnungen Test] ")
            if summary is None:
                self.log_print("[Rechnungen Test] No Rechnungen data detected.")
                return
            self._log_rechnungen_summary("[Rechnungen Test] ", summary)
        except Exception as e:
            self.log_print(f"[Rechnungen Test] ERROR: {e!r}")

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

            skip_waits = self._should_skip_manual_waits()
            list_wait = (
                0.0 if skip_waits else float(self.cfg.get("post_search_wait", 1.2))
            )

            results = []
            total = len(queries)
            for idx, (aktenzeichen, _row) in enumerate(queries, 1):
                prefix = f"[Rechnungen {idx}/{total}] "
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

                summary = self._extract_rechnungen_summary(prefix=prefix)
                if summary is None:
                    self.log_print(
                        f"{prefix}Rechnungen capture returned no data; storing defaults."
                    )
                    summary = self._summarize_rechnungen_entries([])
                else:
                    self._log_rechnungen_summary(prefix, summary)

                results.append(
                    self._build_rechnungen_result_row(aktenzeichen, summary)
                )

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

    def run_streitwert_threaded(self):
        t = threading.Thread(target=self.run_streitwert, daemon=True)
        t.start()

    def run_streitwert_with_rechnungen_threaded(self):
        t = threading.Thread(
            target=self.run_streitwert_with_rechnungen, daemon=True
        )
        t.start()

    def run_streitwert_with_rechnungen(self):
        self.run_streitwert(include_rechnungen=True)

    def _filter_streitwert_rows(self, lines):
        inc = [
            t.strip().lower()
            for t in (self.includes_var.get() or "").split(",")
            if t.strip()
        ]
        exc = [
            t.strip().lower()
            for t in (self.excludes_var.get() or "").split(",")
            if t.strip()
        ]
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
            if excl_k and re.match(r"^\s*k", low_raw):
                debug_rows.append((raw, "excluded prefix 'K'"))
                continue
            if exc and any(tok in low_raw or tok in low_norm for tok in exc):
                debug_rows.append((raw, "matched exclude token"))
                continue
            matched_token = None
            if inc:
                for tok in inc:
                    if tok in low_raw or tok in low_norm:
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

    def _capture_rechnungen_lines(self, prefix=""):
        if not self.current_rect:
            self.log_print(f"{prefix}No active RDP rectangle. Connect before capturing.")
            return []
        if "rechnungen_region" not in self.cfg:
            self.log_print(f"{prefix}Rechnungen region is not configured.")
            return []
        try:
            img, scale = _grab_region_color_generic(
                self.current_rect,
                self.cfg["rechnungen_region"],
                self.upscale_var.get(),
            )
        except Exception as exc:
            self.log_print(
                f"{prefix}Failed to capture Rechnungen region: {exc}"
            )
            return []
        df = do_ocr_data(
            img, lang=self.lang_var.get().strip() or "deu+eng", psm=6
        )
        lines = lines_from_tsv(df, scale=scale)
        self.log_print(
            f"{prefix}Rechnungen OCR lines: {len(lines)}."
        )
        return lines

    def _parse_rechnungen_entries(self, lines, prefix=""):
        entries = []
        skipped = []
        for x, y, w, h, text in lines:
            raw = (text or "").strip()
            if not raw:
                continue
            norm = normalize_line(raw)
            amount = extract_amount_from_text(norm)
            amount = clean_amount_display(amount) if amount else None
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
            entry = {
                "raw": raw,
                "norm": norm,
                "amount": amount,
                "date": date_text,
                "date_obj": date_obj,
                "invoice": invoice,
                "x": x,
                "y": y,
                "w": w,
                "h": h,
            }
            entries.append(entry)
        entries.sort(
            key=lambda e: (
                e.get("date_obj") or datetime.min,
                e.get("y", 0),
                e.get("x", 0),
            )
        )
        for entry in entries:
            self.log_print(
                f"{prefix}Rechnungen candidate: {entry['norm']}"
            )
        for norm, reason in skipped[:6]:
            self.log_print(f"{prefix}Skipped Rechnungen line '{norm}' ({reason}).")
        return entries

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
            "gg": _copy(gg_entry) if gg_entry else {
                "amount": "0",
                "date": "",
                "invoice": "",
                "raw": "",
            },
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
            self.log_print(f"{prefix}-Received GG: {gg_entry.get('amount', '')}{detail}")
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
            self.log_print(
                f"{prefix}Focusing doc list at ({focus_x}, {focus_y})."
            )
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
        if not self.current_rect or not rel_box:
            return None
        try:
            rx, ry, rw, rh = rel_to_abs(self.current_rect, rel_box)
        except Exception:
            return None
        try:
            img, scale = _grab_region_color_generic(
                self.current_rect, rel_box, self.upscale_var.get()
            )
        except Exception:
            return None
        df = do_ocr_data(
            img, lang=self.lang_var.get().strip() or "deu+eng", psm=6
        )
        lines = lines_from_tsv(df, scale=scale)
        overlay = self._find_overlay_entry(lines)
        if not overlay:
            return None
        entry = overlay.copy()
        entry["abs_x"] = rx + overlay["x"]
        entry["abs_y"] = ry + overlay["y"]
        entry["abs_w"] = overlay["w"]
        entry["abs_h"] = overlay["h"]
        return entry

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
                    self.log_print(
                        f"{prefix}Document list overlay cleared{suffix}."
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
                    self.log_print(
                        f"{prefix}Deal search overlay cleared{suffix}."
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

    def _click_pdf_result_button(self, prefix=""):
        if not self.current_rect:
            return False
        point = self.cfg.get("pdf_hits_point")
        if not (isinstance(point, (list, tuple)) and len(point) == 2):
            msg = "PDF results button point is not configured. Please calibrate it."
            self.log_print(f"{prefix}{msg}" if prefix else msg)
            return False
        try:
            Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        except Exception:
            pass
        hx, hy = rel_to_abs(self.current_rect, point)
        self.log_print(f"{prefix}Clicking PDF result button at ({hx}, {hy}).")
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

    def _process_open_pdf(self, prefix="", search_term=None, retype=False):
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

        clicked_hits = self._click_pdf_result_button(prefix=prefix)
        if not clicked_hits:
            self.log_print(
                f"{prefix}Skipped PDF results click; proceeding directly to page OCR."
            )
        else:
            try:
                extra_wait = float(self.cfg.get("pdf_view_extra_wait", 2.0))
            except Exception:
                extra_wait = 2.0
            if extra_wait > 0:
                self.log_print(
                    f"{prefix}Waiting {extra_wait:.1f}s after PDF results click before checking overlays."
                )
                time.sleep(extra_wait)

        reason = "after PDF results click" if clicked_hits else "before page OCR"
        self._wait_for_pdf_ready(prefix=prefix, reason=reason)

        page_img, page_scale = _grab_region_color_generic(
            self.current_rect,
            self.cfg["pdf_text_region"],
            self.upscale_var.get(),
        )
        dft = do_ocr_data(
            page_img, lang=self.lang_var.get().strip() or "deu+eng", psm=6
        )
        lines_pg = lines_from_tsv(dft, scale=page_scale)
        self.log_print(
            f"{prefix}Page OCR lines captured: {len(lines_pg)}. Extracting amount."
        )
        combined = "\n".join(normalize_line(t) for _, _, _, _, t in lines_pg)
        amt = extract_amount_from_text(combined)
        return clean_amount_display(amt) if amt else None

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
            if not queries:
                return

            skip_waits = self._should_skip_manual_waits()
            list_wait = (
                0.0 if skip_waits else float(self.cfg.get("post_search_wait", 1.2))
            )
            doc_wait = (
                0.0 if skip_waits else float(self.docwait_var.get() or 1.2)
            )
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
                    rechnungen_results.append(
                        self._build_rechnungen_result_row(aktenzeichen, rechn_summary)
                    )

                term = self.streitwort_var.get().strip() or "Streitwert"
                self._wait_for_doc_search_ready(
                    prefix=prefix, reason="before PDF search"
                )
                if not self._type_pdf_search(term, prefix=prefix):
                    self.log_print(
                        f"{prefix}Unable to type Streitwert term in the PDF search box."
                    )
                    continue
                self.log_print(
                    f"{prefix}Typed '{term}' into the PDF search box."
                )
                if list_wait > 0:
                    time.sleep(list_wait)
                self._wait_for_doc_search_ready(
                    prefix=prefix, reason="after PDF search"
                )
                self._wait_for_doclist_ready(
                    prefix=prefix, reason="after PDF search"
                )

                rx, ry, rw, rh = doc_rect
                focus_x = rx + max(5, rw // 40)
                focus_y = ry + max(5, rh // 40)
                self.log_print(
                    f"{prefix}Clicking doc list to ensure focus at ({focus_x}, {focus_y})."
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

                if not ordered:
                    reason = ", ".join(
                        f"{r}: {raw}" for raw, r in debug_rows[:4]
                    )
                    if not reason:
                        sample = ", ".join((txt or "").strip() for *_, txt in lines[:4])
                        reason = sample or "no OCR rows"
                    self.log_print(
                        f"{prefix}No matching rows for '{aktenzeichen}'. Details: {reason}"
                    )
                    continue

                first = ordered[0]
                tag = first.get("token") or "any"
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

                self.log_print(
                    f"{prefix}Clicked View button for the selected row."
                )

                if doc_wait > 0:
                    time.sleep(doc_wait)
                amount = self._process_open_pdf(prefix=prefix)
                results.append(
                    {
                        "aktenzeichen": aktenzeichen,
                        "row_text": first["norm"],
                        "amount": amount or "",
                    }
                )
                self.simple_log_print(
                    f"{aktenzeichen}: {amount or '(none)'}"
                )
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
                self.log_print("No Streitwert results were collected from the Excel list.")

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
            if not self._type_pdf_search(term, prefix="[Test] "):
                self.log_print("[Test] Unable to type the Streitwert search term.")
                return
            self.log_print(f"[Test] Typed '{term}' into the PDF search box.")
            if list_wait > 0:
                time.sleep(list_wait)
            self._wait_for_doc_search_ready(
                prefix="[Test] ", reason="after PDF search"
            )
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

            self.log_print(
                f"[Test] Doc list OCR lines: {len(lines)} | includes: {inc or ['(none)']} | excludes: {exc or ['(none)']}"
            )
            if not ordered:
                preview = debug_rows[:5] or [(raw, "") for *_, raw in lines[:5]]
                for raw, reason in preview:
                    desc = f"  {reason or 'OCR'} → {raw}"
                    self.log_print(desc)
                self.log_print("[Test] No rows matched the include tokens after typing 'Streitwert'.")
                return

            first = ordered[0]
            tag = first.get("token") or "any"
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
                prefix="[Test] ", search_term=term or "Streitwert"
            )
            self.log_print(
                f"[Test] Extracted Streitwert amount: {amount or '(none)'}"
            )

            self._close_active_pdf(prefix="[Test] ")
            time.sleep(0.5)
            self.log_print("[Test] Closed PDF after verification.")

            self.log_print("[Test] Streitwert setup check finished.")

        except Exception as e:
            self.log_print("ERROR during Streitwert test: " + repr(e))

    # ---------- Utilities ----------
    def show_preview(self, img: Image.Image):
        preview = img.copy()
        preview.thumbnail((720, 240))
        self.ocr_preview_imgtk = ImageTk.PhotoImage(preview)
        self.img_label.configure(image=self.ocr_preview_imgtk)

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
            cap_left = int(
                min(max(cx - cap_w // 2, left), max(left, right - cap_w))
            )
            cap_top = int(
                min(max(cy - cap_h // 2, top), max(top, bottom - cap_h))
            )
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
        draw.rectangle([(0, 0), (img.width - 1, img.height - 1)], outline="yellow", width=2)

        info_text = "cursor outside"
        if cursor_inside and cx is not None and cy is not None:
            local_x = cx - cap_left
            local_y = cy - cap_top
            draw.line([(local_x - 12, local_y), (local_x + 12, local_y)], fill="red", width=2)
            draw.line([(local_x, local_y - 12), (local_x, local_y + 12)], fill="red", width=2)
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

    def _should_skip_manual_waits(self):
        try:
            return bool(self.skip_waits_var.get())
        except Exception:
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

        self.cfg["includes"] = self.includes_var.get().strip()
        self.cfg["excludes"] = self.excludes_var.get().strip()
        self.cfg["exclude_prefix_k"] = bool(self.exclude_k_var.get())
        self.cfg["streitwert_term"] = (
            self.streitwort_var.get().strip() or "Streitwert"
        )
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
        self.cfg["streitwert_overlay_skip_waits"] = bool(
            self.skip_waits_var.get()
        )
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
            self.streitwort_var.set(
                self.cfg.get("streitwert_term", "Streitwert")
            )
            self.docwait_var.set(str(self.cfg.get("doc_open_wait", 1.2)))
            self.hitwait_var.set(str(self.cfg.get("pdf_hit_wait", 1.0)))
            self.streit_csv_var.set(
                self.cfg.get(
                    "streitwert_results_csv", "streitwert_results.csv"
                )
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
            if hasattr(self, "skip_waits_var"):
                self.skip_waits_var.set(
                    self.cfg.get("streitwert_overlay_skip_waits", False)
                )
            if hasattr(self, "pdf_view_wait_var"):
                self.pdf_view_wait_var.set(
                    str(self.cfg.get("pdf_view_extra_wait", 2.0))
                )
            hits_pt = self.cfg.get("pdf_hits_point")
            if not (
                isinstance(hits_pt, (list, tuple)) and len(hits_pt) == 2
            ):
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
