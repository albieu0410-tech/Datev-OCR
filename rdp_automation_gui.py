import os, io, json, time, threading, re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk, ImageFilter, ImageOps, ImageStat
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
    # -------- NEW: Rechnungen capture --------
    "rechnungen_region": [0.0, 0.0, 0.0, 0.0],
    "rechnungen_csv": "Streitwert_Results_Rechnungen.csv",
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
        cfg.setdefault("rechnungen_region", [0.0, 0.0, 0.0, 0.0])
        cfg.setdefault("rechnungen_csv", DEFAULTS["rechnungen_csv"])
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


# ---------- Normalization / parsing ----------
AMOUNT_RE = re.compile(r"\b\d+(?:[.\s]\d{3})*[.,]\d{2}\s*(?:EUR|€)?\b", re.IGNORECASE)
DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
INVOICE_RE = re.compile(r"\b\d{5,}\b")
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


def lines_from_tsv(tsv_df):
    if tsv_df is None or tsv_df.empty:
        return []
    df = tsv_df.dropna(subset=["text"])
    df = df[df["conf"] > -1]
    lines = []
    for (_, _, _), grp in df.groupby(["block_num", "par_num", "line_num"]):
        ys = grp["top"].min()
        txt = " ".join(str(t) for t in grp["text"] if str(t).strip())
        if txt.strip():
            lines.append((int(ys), txt.strip()))
    lines.sort(key=lambda x: x[0])
    return lines


def extract_amount_from_lines(lines, keyword=None):
    if not lines:
        return None, None

    norm_lines = [(y, normalize_line_soft(t)) for y, t in lines]

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
        self._run_rechnungen_first = False

        self.create_widgets()

    def create_widgets(self):
        # --- Left frame (controls) ---
        left = ttk.Frame(self, padding=10)
        left.pack(side=tk.LEFT, fill=tk.Y)

        # RDP
        ttk.Label(left, text="RDP Window Title (Regex)").pack(anchor="w")
        self.rdp_var = tk.StringVar(value=self.cfg["rdp_title_regex"])
        ttk.Entry(left, textvariable=self.rdp_var, width=52).pack(
            anchor="w", pady=(0, 6)
        )

        # Excel path
        ttk.Label(left, text="Excel Path").pack(anchor="w")
        xframe = ttk.Frame(left)
        xframe.pack(anchor="w", fill=tk.X)
        self.xls_var = tk.StringVar(value=self.cfg["excel_path"])
        ttk.Entry(xframe, textvariable=self.xls_var, width=42).pack(
            side=tk.LEFT, pady=2
        )
        ttk.Button(xframe, text="Browse", command=self.browse_excel).pack(
            side=tk.LEFT, padx=6
        )

        # Sheet / Start cell / Max rows
        row1 = ttk.Frame(left)
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

        # Fallback input column (if start cell is empty)
        ttk.Label(left, text="(Optional) Input column name (if Start cell empty)").pack(
            anchor="w", pady=(6, 0)
        )
        self.col_var = tk.StringVar(value=self.cfg["input_column"])
        ttk.Entry(left, textvariable=self.col_var, width=20).pack(
            anchor="w", pady=(0, 6)
        )

        # Tesseract path + lang
        ttk.Label(left, text="Tesseract Path (exe or folder)").pack(anchor="w")
        tframe = ttk.Frame(left)
        tframe.pack(anchor="w", fill=tk.X)
        self.tess_var = tk.StringVar(value=self.cfg["tesseract_path"])
        ttk.Entry(tframe, textvariable=self.tess_var, width=42).pack(
            side=tk.LEFT, pady=2
        )
        ttk.Button(tframe, text="Browse", command=self.browse_tesseract).pack(
            side=tk.LEFT, padx=6
        )

        ttk.Label(left, text="OCR language (e.g., deu+eng)").pack(
            anchor="w", pady=(6, 0)
        )
        self.lang_var = tk.StringVar(value=self.cfg.get("tesseract_lang", "deu+eng"))
        ttk.Entry(left, textvariable=self.lang_var, width=16).pack(
            anchor="w", pady=(0, 6)
        )

        ttk.Separator(left).pack(fill=tk.X, pady=8)

        # Timing
        r1 = ttk.Frame(left)
        r1.pack(anchor="w")
        ttk.Label(r1, text="Typing delay (sec/char)").pack(side=tk.LEFT)
        self.type_var = tk.StringVar(value=str(self.cfg["type_delay"]))
        ttk.Entry(r1, textvariable=self.type_var, width=8).pack(side=tk.LEFT, padx=6)

        r2 = ttk.Frame(left)
        r2.pack(anchor="w", pady=(4, 0))
        ttk.Label(r2, text="Post-search wait (sec)").pack(side=tk.LEFT)
        self.wait_var = tk.StringVar(value=str(self.cfg["post_search_wait"]))
        ttk.Entry(r2, textvariable=self.wait_var, width=8).pack(side=tk.LEFT, padx=6)

        # Typing test
        r3 = ttk.Frame(left)
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

        ttk.Separator(left).pack(fill=tk.X, pady=8)

        # Calibration
        ttk.Label(left, text="Calibration").pack(anchor="w")
        cframe = ttk.Frame(left)
        cframe.pack(anchor="w")
        ttk.Button(cframe, text="Connect RDP", command=self.connect_rdp).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(
            cframe, text="Pick Search Point", command=self.pick_search_point
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            cframe, text="Pick Result Region", command=self.pick_result_region
        ).pack(side=tk.LEFT, padx=2)

        ttk.Label(left, text="Search Point (x%, y%)").pack(anchor="w", pady=(8, 0))
        self.sp_var = tk.StringVar(
            value=f"{self.cfg['search_point'][0]:.3f}, {self.cfg['search_point'][1]:.3f}"
        )
        ttk.Entry(left, textvariable=self.sp_var, width=30).pack(anchor="w")

        ttk.Label(left, text="Result Region (l%, t%, w%, h%)").pack(
            anchor="w", pady=(8, 0)
        )
        rr = self.cfg["result_region"]
        self.rr_var = tk.StringVar(
            value=f"{rr[0]:.3f}, {rr[1]:.3f}, {rr[2]:.3f}, {rr[3]:.3f}"
        )
        ttk.Entry(left, textvariable=self.rr_var, width=40).pack(anchor="w")

        # OCR options & full-region parsing
        rowb = ttk.Frame(left)
        rowb.pack(anchor="w", pady=(6, 0))
        ttk.Label(rowb, text="Upscale ×").pack(side=tk.LEFT)
        self.upscale_var = tk.StringVar(value=str(self.cfg.get("upscale_x", 4)))
        ttk.Entry(rowb, textvariable=self.upscale_var, width=5).pack(
            side=tk.LEFT, padx=6
        )
        self.color_var = tk.BooleanVar(value=self.cfg.get("color_ocr", True))
        ttk.Checkbutton(rowb, text="Color OCR", variable=self.color_var).pack(
            side=tk.LEFT, padx=6
        )

        fr = ttk.Frame(left)
        fr.pack(anchor="w", pady=(8, 0))
        self.fullparse_var = tk.BooleanVar(
            value=self.cfg.get("use_full_region_parse", True)
        )
        ttk.Checkbutton(
            fr, text="Use full-region parsing", variable=self.fullparse_var
        ).pack(side=tk.LEFT)
        ttk.Label(fr, text="Keyword").pack(side=tk.LEFT, padx=(12, 4))
        self.keyword_var = tk.StringVar(value=self.cfg.get("keyword", "Honorar"))
        ttk.Entry(fr, textvariable=self.keyword_var, width=16).pack(side=tk.LEFT)

        nr = ttk.Frame(left)
        nr.pack(anchor="w", pady=(6, 0))
        self.normalize_var = tk.BooleanVar(value=self.cfg.get("normalize_ocr", True))
        ttk.Checkbutton(
            nr, text="Normalize OCR (O→0, S→5…)", variable=self.normalize_var
        ).pack(side=tk.LEFT)

        # ---------------- Tabs: Streitwert & Rechnungen ----------------
        tabs = ttk.Notebook(left)
        tabs.pack(anchor="w", fill=tk.X, pady=10)

        self.streitwert_tab = ttk.Frame(tabs, padding=6)
        self.rechnungen_tab = ttk.Frame(tabs, padding=6)
        tabs.add(self.streitwert_tab, text="Streitwert")
        tabs.add(self.rechnungen_tab, text="Rechnungen")

        # Streitwert tab contents
        ttk.Label(self.streitwert_tab, text="Results CSV").pack(anchor="w")
        self.csv_var = tk.StringVar(value=self.cfg["results_csv"])
        ttk.Entry(
            self.streitwert_tab, textvariable=self.csv_var, width=42
        ).pack(anchor="w", pady=(0, 4))

        ttk.Label(self.streitwert_tab, text="Rechnungen CSV").pack(
            anchor="w", pady=(4, 0)
        )
        self.rechnungen_csv_var = tk.StringVar(
            value=self.cfg.get("rechnungen_csv", DEFAULTS["rechnungen_csv"])
        )
        ttk.Entry(
            self.streitwert_tab, textvariable=self.rechnungen_csv_var, width=42
        ).pack(anchor="w", pady=(0, 4))

        ttk.Button(
            self.streitwert_tab,
            text="Start Streitwert Scan + Rechnungen",
            command=self.run_streitwert_with_rechnungen_threaded,
        ).pack(anchor="w", pady=(8, 0))

        # Rechnungen tab contents
        ttk.Label(
            self.rechnungen_tab,
            text="Rechnungen Region (l%, t%, w%, h%)",
        ).pack(anchor="w")
        rr = self.cfg.get("rechnungen_region", [0.0, 0.0, 0.0, 0.0])
        self.rechnungen_region_var = tk.StringVar(
            value=", ".join(f"{float(v):.3f}" for v in rr)
        )
        ttk.Entry(
            self.rechnungen_tab, textvariable=self.rechnungen_region_var, width=40
        ).pack(anchor="w", pady=(0, 4))

        rbtn = ttk.Frame(self.rechnungen_tab)
        rbtn.pack(anchor="w", pady=(4, 0))
        ttk.Button(
            rbtn, text="Pick Rechnungen Region", command=self.pick_rechnungen_region
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            rbtn, text="Test Rechnungen Extraction", command=self.test_rechnungen
        ).pack(side=tk.LEFT, padx=2)

        ttk.Label(
            self.rechnungen_tab,
            text="Use the test button to verify OCR before running the scan.",
        ).pack(anchor="w", pady=(6, 0))

        ttk.Button(
            left, text="Test Parse (full region)", command=self.test_parse_full
        ).pack(anchor="w", pady=(8, 0))

        # ---------------- Amount Region Profiles (NEW) ----------------
        ttk.Separator(left).pack(fill=tk.X, pady=10)
        ttk.Label(left, text="Amount Region Profiles").pack(anchor="w")

        prof_row1 = ttk.Frame(left)
        prof_row1.pack(anchor="w", pady=(4, 0))
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

        prof_row2 = ttk.Frame(left)
        prof_row2.pack(anchor="w", pady=(6, 0))
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

        prof_row3 = ttk.Frame(left)
        prof_row3.pack(anchor="w", pady=(6, 0))
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

        # Save/Load + Run
        tframe2 = ttk.Frame(left)
        tframe2.pack(anchor="w", pady=8)
        ttk.Button(tframe2, text="Save Config", command=self.save_config).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(tframe2, text="Load Config", command=self.load_config).pack(
            side=tk.LEFT, padx=2
        )

        ttk.Separator(left).pack(fill=tk.X, pady=8)
        ttk.Button(left, text="Run Batch", command=self.run_batch_threaded).pack(
            anchor="w", pady=(0, 6)
        )

        # --- Right frame (preview + log) ---
        right = ttk.Frame(self, padding=10)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(right, text="OCR Preview (color crop)").pack(anchor="w")
        self.img_label = ttk.Label(right)
        self.img_label.pack(anchor="w", pady=(0, 6))
        ttk.Label(right, text="Log").pack(anchor="w")
        self.log = tk.Text(right, height=22)
        self.log.pack(fill=tk.BOTH, expand=True)

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

    def pick_search_point(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        messagebox.showinfo(
            "Pick Search Point",
            "Position your mouse over the search bar location.\nWaiting 3 seconds after you click OK...",
        )
        self.update()
        time.sleep(3)
        x, y = pyautogui.position()
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
        messagebox.showinfo(
            "Pick Result Region",
            "Position your mouse over the TOP-LEFT corner of the region.\nWaiting 3 seconds after you click OK...",
        )
        self.update()
        time.sleep(3)
        x1, y1 = pyautogui.position()

        messagebox.showinfo(
            "Pick Result Region",
            "Now position your mouse over the BOTTOM-RIGHT corner.\nWaiting 3 seconds after you click OK...",
        )
        self.update()
        time.sleep(3)
        x2, y2 = pyautogui.position()
        left, top = min(x1, x2), min(y1, y2)
        width, height = abs(x2 - x1), abs(y2 - y1)
        rel_box = abs_to_rel(self.current_rect, abs_box=(left, top, width, height))
        self.cfg["result_region"] = rel_box
        self.rr_var.set(
            f"{rel_box[0]:.3f}, {rel_box[1]:.3f}, {rel_box[2]:.3f}, {rel_box[3]:.3f}"
        )
        self.log_print(f"Result region set (relative): {rel_box}")

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
        messagebox.showinfo(
            "Pick Amount Region",
            "Place your mouse at the TOP-LEFT of the amount area inside the Result Region.\nWaiting 3 seconds...",
        )
        self.update()
        time.sleep(3)
        x1, y1 = pyautogui.position()

        messagebox.showinfo(
            "Pick Amount Region", "Now the BOTTOM-RIGHT.\nWaiting 3 seconds..."
        )
        self.update()
        time.sleep(3)
        x2, y2 = pyautogui.position()

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

    def pick_rechnungen_region(self):
        if not self.current_rect:
            self.connect_rdp()
            if not self.current_rect:
                return
        Desktop(backend="uia").window(title_re=self.rdp_var.get()).set_focus()
        messagebox.showinfo(
            "Pick Rechnungen Region",
            "Position your mouse over the TOP-LEFT corner of the Rechnungen area.\n"
            "Waiting 3 seconds after you click OK...",
        )
        self.update()
        time.sleep(3)
        x1, y1 = pyautogui.position()

        messagebox.showinfo(
            "Pick Rechnungen Region",
            "Now position your mouse over the BOTTOM-RIGHT corner.\nWaiting 3 seconds after you click OK...",
        )
        self.update()
        time.sleep(3)
        x2, y2 = pyautogui.position()

        left, top = min(x1, x2), min(y1, y2)
        width, height = abs(x2 - x1), abs(y2 - y1)
        rel_box = abs_to_rel(self.current_rect, abs_box=(left, top, width, height))
        self.cfg["rechnungen_region"] = rel_box
        if hasattr(self, "rechnungen_region_var"):
            self.rechnungen_region_var.set(
                ", ".join(f"{v:.3f}" for v in rel_box)
            )
        self.log_print(
            f"Rechnungen region set (relative): {', '.join(f'{v:.3f}' for v in rel_box)}"
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

    def test_rechnungen(self):
        try:
            self.apply_paths_to_tesseract()
            if not self.current_rect:
                self.connect_rdp()
                if not self.current_rect:
                    return
            crop, lines, entries, summary = self._read_rechnungen_region()
            self.show_preview(crop)
            if lines:
                joined = "\n".join(t for _, t in lines if t.strip())
                self.log_print("Rechnungen OCR:\n" + joined)
            else:
                self.log_print("Rechnungen OCR:\n(no text)")
            self._log_rechnungen_summary(summary)
        except Exception as e:
            messagebox.showerror("Rechnungen Test", f"Failed: {e}")

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
        full_text = "\n".join(t for _, t in lines)
        if self.normalize_var.get():
            lines = [(y, normalize_line_soft(t)) for y, t in lines]
            full_text = "\n".join(t for _, t in lines)
        amount, line = extract_amount_from_lines(lines, keyword=keyword)
        return full_text, crop, lines, amount

    # ---------- Rechnungen helpers ----------
    def _validate_rechnungen_region(self):
        l, t, w, h = self.cfg.get("rechnungen_region", [0.0, 0.0, 0.0, 0.0])
        if w <= 0 or h <= 0:
            raise ValueError(
                "Invalid Rechnungen Region. Please pick a non-zero area in the Rechnungen tab."
            )

    def _grab_rechnungen_region(self):
        self._validate_rechnungen_region()
        rx, ry, rw, rh = rel_to_abs(self.current_rect, self.cfg["rechnungen_region"])
        region = grab_xywh(rx, ry, rw, rh)
        scale = max(1, int(self.upscale_var.get() or 3))
        return upscale_pil(region, scale=scale)

    def _amount_to_float(self, amount_text: str) -> float:
        clean = amount_text.upper().replace("EUR", "").replace("€", "")
        clean = clean.replace(" ", "")
        clean = clean.replace(".", "").replace(",", ".")
        return float(clean)

    def _format_currency(self, value: float) -> str:
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _parse_rechnungen_entries(self, lines):
        entries = []
        seen = set()
        for _, raw in lines:
            text = str(raw or "").strip()
            if not text:
                continue
            text = text.translate(DIGIT_FIX)
            text = re.sub(r"\s+", " ", text)
            date_match = DATE_RE.search(text)
            amount_match = AMOUNT_RE.search(text)
            if not date_match or not amount_match:
                continue
            date_str = date_match.group(0)
            amount_text = amount_match.group(0)
            trailing = text[amount_match.end() :]
            invoice_match = INVOICE_RE.search(trailing)
            invoice = ""
            if invoice_match:
                invoice = re.sub(r"\D", "", invoice_match.group(0))
            try:
                amount_value = self._amount_to_float(amount_text)
            except ValueError:
                continue
            try:
                date_val = datetime.strptime(date_str, "%d.%m.%Y")
            except ValueError:
                continue
            key = (date_str, round(amount_value, 2), invoice)
            if key in seen:
                continue
            seen.add(key)
            entries.append(
                {
                    "date": date_val,
                    "date_str": date_str,
                    "amount": amount_value,
                    "amount_str": self._format_currency(amount_value),
                    "invoice": invoice,
                    "raw": text,
                }
            )
        return entries

    def _summarize_rechnungen_entries(self, entries):
        total_candidates = [e for e in entries if not e.get("invoice")]
        total_entry = (
            max(total_candidates, key=lambda e: e["date"]) if total_candidates else None
        )

        invoice_entries = [e for e in entries if e.get("invoice")]
        invoice_entries.sort(key=lambda e: e["date"])
        if invoice_entries:
            received_court = invoice_entries[-1]
            received_gg = invoice_entries[0] if len(invoice_entries) > 1 else None
        else:
            received_court = None
            received_gg = None

        return {
            "total": total_entry,
            "received_court": received_court,
            "received_gg": received_gg,
        }

    def _read_rechnungen_region(self):
        crop = self._grab_rechnungen_region()
        lang = self.lang_var.get().strip() or "deu+eng"
        df = do_ocr_data(crop, lang=lang, psm=6)
        lines = lines_from_tsv(df)
        cleaned = []
        for y, text in lines:
            txt = str(text or "").strip()
            if not txt:
                continue
            txt = txt.translate(DIGIT_FIX)
            txt = re.sub(r"\s+", " ", txt)
            cleaned.append((y, txt))
        cleaned.sort(key=lambda x: x[0])
        entries = self._parse_rechnungen_entries(cleaned)
        summary = self._summarize_rechnungen_entries(entries)
        return crop, cleaned, entries, summary

    def _summary_amount_text(self, entry):
        if not entry:
            return "0,00 EUR"
        return f"{entry['amount_str']} EUR"

    def _log_rechnungen_summary(self, summary):
        total_text = self._summary_amount_text(summary.get("total"))
        self.log_print(f"-Total Fees: {total_text}")

        court_entry = summary.get("received_court")
        if court_entry:
            details = court_entry["date_str"]
            if court_entry.get("invoice"):
                details += f" {court_entry['invoice']}"
            self.log_print(
                f"-Received Court Fees ({details}) {self._summary_amount_text(court_entry)}"
            )
        else:
            self.log_print("-Received Court Fees: 0,00 EUR")

        gg_entry = summary.get("received_gg")
        if gg_entry:
            details = gg_entry["date_str"]
            if gg_entry.get("invoice"):
                details += f" {gg_entry['invoice']}"
            self.log_print(
                f"-Received GG ({details}) {self._summary_amount_text(gg_entry)}"
            )
        else:
            self.log_print("-Received GG: 0,00 EUR")

    def _save_rechnungen_csv(self, summary, lines):
        path = (self.rechnungen_csv_var.get() or "").strip() or DEFAULTS["rechnungen_csv"]
        total = summary.get("total")
        court = summary.get("received_court")
        gg = summary.get("received_gg")
        row = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "total_fees": self._summary_amount_text(total),
            "total_fees_date": total["date_str"] if total else "",
            "total_fees_invoice": total.get("invoice", "") if total else "",
            "received_court_fees": self._summary_amount_text(court),
            "received_court_fees_date": court["date_str"] if court else "",
            "received_court_fees_invoice": court.get("invoice", "") if court else "",
            "received_gg": self._summary_amount_text(gg),
            "received_gg_date": gg["date_str"] if gg else "",
            "received_gg_invoice": gg.get("invoice", "") if gg else "",
            "raw_ocr": " | ".join(t for _, t in (lines or [])),
        }
        pd.DataFrame([row]).to_csv(path, index=False, encoding="utf-8-sig")
        self.log_print(f"Rechnungen summary saved to {path}")

    # ---------- Batch ----------
    def run_streitwert_with_rechnungen_threaded(self):
        self._run_rechnungen_first = True
        self.run_batch_threaded()

    def run_batch_threaded(self):
        threading.Thread(target=self.run_batch, daemon=True).start()

    def run_batch(self):
        try:
            self.pull_form_into_cfg()
            save_cfg(self.cfg)
            self.apply_paths_to_tesseract()

            _, rect = connect_rdp_window(self.rdp_var.get())
            self.current_rect = rect

            run_rechnungen = self._run_rechnungen_first
            self._run_rechnungen_first = False
            if run_rechnungen:
                try:
                    crop, lines, entries, summary = self._read_rechnungen_region()
                    self._log_rechnungen_summary(summary)
                    self._save_rechnungen_csv(summary, lines)
                except Exception as e:
                    self.log_print(f"ERROR during Rechnungen extraction: {e}")

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

    # ---------- Utilities ----------
    def show_preview(self, img: Image.Image):
        preview = img.copy()
        preview.thumbnail((720, 240))
        self.ocr_preview_imgtk = ImageTk.PhotoImage(preview)
        self.img_label.configure(image=self.ocr_preview_imgtk)

    def log_print(self, text):
        self.log.insert(tk.END, str(text) + "\n")
        self.log.see(tk.END)
        self.update_idletasks()

    def pull_form_into_cfg(self):
        self.cfg["rdp_title_regex"] = self.rdp_var.get().strip()
        self.cfg["excel_path"] = self.xls_var.get().strip()
        sv = self.sheet_var.get().strip()
        self.cfg["excel_sheet"] = int(sv) if sv.isdigit() else sv
        self.cfg["input_column"] = self.col_var.get().strip()
        self.cfg["results_csv"] = self.csv_var.get().strip()
        self.cfg["rechnungen_csv"] = (
            (self.rechnungen_csv_var.get() or "").strip()
            or DEFAULTS["rechnungen_csv"]
        )
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

        region_text = (self.rechnungen_region_var.get() or "").strip()
        parts = [p.strip() for p in region_text.split(",")]
        if len(parts) == 4:
            try:
                self.cfg["rechnungen_region"] = [
                    float(p.replace("%", "")) for p in parts
                ]
            except ValueError:
                pass

    def load_config(self):
        try:
            self.cfg = load_cfg()  # Load from file

            # Update form values from cfg
            self.rdp_var.set(self.cfg["rdp_title_regex"])
            self.xls_var.set(self.cfg["excel_path"])
            self.sheet_var.set(str(self.cfg["excel_sheet"]))
            self.col_var.set(self.cfg["input_column"])
            self.csv_var.set(self.cfg["results_csv"])
            self.rechnungen_csv_var.set(
                self.cfg.get("rechnungen_csv", DEFAULTS["rechnungen_csv"])
            )
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

            rr = self.cfg.get("rechnungen_region", [0.0, 0.0, 0.0, 0.0])
            self.rechnungen_region_var.set(
                ", ".join(f"{float(v):.3f}" for v in rr)
            )

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
