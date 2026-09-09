"""
PDF2text - Desktop Document & OCR Converter
Converts PDF, scanned documents, and images to Excel (.xlsx), Word (.docx), and Text (.txt).
Powered by high-speed RapidOCR (ONNX Runtime), PyMuPDF zero-copy rendering,
spatial table extraction with category row merging, and Egyptian construction dictionary.
"""

import os
import re
import sys
import json
import time
import threading
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Ensure UTF-8 output on Windows console
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Optional CustomTkinter for modern UI styling
try:
    import customtkinter as ctk
    HAS_CTK = True
except ImportError:
    HAS_CTK = False

# High-speed document, computer vision & OCR processing libraries
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

try:
    import pymupdf
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    from rapidocr_onnxruntime import RapidOCR
    from rapidocr_onnxruntime.rapid_ocr_api import read_yaml, root_dir as rapidocr_root_dir
    from rapidocr_onnxruntime.ch_ppocr_v3_rec import TextRecognizer
    HAS_RAPIDOCR = True
except ImportError:
    HAS_RAPIDOCR = False

try:
    import rapidfuzz
    from rapidfuzz import process, fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import docx
    from docx.shared import Pt, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

FORMAT_OPTIONS = [
    "Excel (.xlsx)",
    "Word (.docx)",
    "Text (.txt)"
]

FORMAT_CONFIG = {
    "Excel (.xlsx)": {
        "ext": ".xlsx",
        "name": "Excel Spreadsheet",
        "filetypes": [("Excel Workbook (*.xlsx)", "*.xlsx"), ("All Files (*.*)", "*.*")],
        "desc": "Extracts multi-column tables with aligned headers, rows, and cells into distinct Excel columns."
    },
    "Word (.docx)": {
        "ext": ".docx",
        "name": "Word Document",
        "filetypes": [("Word Document (*.docx)", "*.docx"), ("All Files (*.*)", "*.*")],
        "desc": "Extracts structured Word document with native Word tables, page headings, and cell formatting."
    },
    "Text (.txt)": {
        "ext": ".txt",
        "name": "Text Document",
        "filetypes": [("Text Document (*.txt)", "*.txt"), ("All Files (*.*)", "*.*")],
        "desc": "Extracts tab-delimited clean UTF-8 text with page headers and table row alignment."
    }
}


def get_base_dir() -> str:
    """Returns application base directory whether frozen EXE or source script."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_model_storage_dir() -> str:
    """Locates offline bundled RapidOCR model weights (arabic_rec.onnx, arabic_dict.txt)."""
    base_dir = get_base_dir()
    candidates = [
        os.path.join(base_dir, "models"),
        os.path.join(os.path.dirname(sys.executable), "models") if getattr(sys, "frozen", False) else None,
        os.path.abspath("models")
    ]
    for p in candidates:
        if p and os.path.isdir(p):
            return p
    return os.path.join(base_dir, "models")


def get_dictionary_path() -> str:
    """Locates Egyptian construction dictionary (dictionary.json)."""
    base_dir = get_base_dir()
    candidates = [
        os.path.join(base_dir, "dictionary.json"),
        os.path.join(os.path.dirname(sys.executable), "dictionary.json") if getattr(sys, "frozen", False) else None,
        os.path.abspath("dictionary.json")
    ]
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return os.path.join(base_dir, "dictionary.json")


def create_rapidocr_engine(models_dir: str = None):
    """Initializes high-speed RapidOCR instance configured with Arabic ONNX model."""
    if not HAS_RAPIDOCR:
        raise RuntimeError("RapidOCR is not installed. Please run 'pip install rapidocr_onnxruntime onnxruntime'.")

    if models_dir is None:
        models_dir = get_model_storage_dir()

    rec_path = os.path.join(models_dir, "arabic_rec.onnx")
    dict_path = os.path.join(models_dir, "arabic_dict.txt")

    engine = RapidOCR()
    if os.path.isfile(rec_path) and os.path.isfile(dict_path):
        config = read_yaml(str(rapidocr_root_dir / "config.yaml"))["Rec"]
        config["model_path"] = rec_path
        config["keys_path"] = dict_path
        engine.text_recognizer = TextRecognizer(config)
    return engine


def normalize_rapidocr_arabic(raw_text: str) -> str:
    """
    Normalizes raw text from RapidOCR Arabic recognition:
    - PaddleOCR/RapidOCR Arabic reads characters in visual left-to-right order.
    - Reversing the string restores natural Arabic reading order.
    - Re-inverts numbers and Latin/English symbols to maintain natural LTR order.
    """
    if not raw_text:
        return ""
    t = raw_text.strip()
    if not any("\u0600" <= ch <= "\u06FF" for ch in t):
        return t

    rev = t[::-1]
    # Restore numbers, dimensions, and English acronyms to LTR
    fixed = re.sub(r"([A-Za-z0-9\-_./*×+]+)", lambda m: m.group(1)[::-1], rev)
    # Format dimensions cleanly (e.g. 60*90 سم)
    fixed = re.sub(r"(\d+)\s*['’`\"*xX×*٢،,]\s*(\d+)\s*(?:س?م)?", r"\1*\2 سم", fixed)
    return StructuredTableExtractor.clean_arabic_dimensions(fixed)


# ==============================================================================
# Arabic Construction Dictionary & Fuzzy Spell Correction
# ==============================================================================
class ArabicDictionaryCorrector:
    """
    Egyptian Arabic construction dictionary and fuzzy OCR spell-correction engine.
    Corrects OCR letter misreadings, typos, and standardizes engineering terminology and units.
    """
    _instance = None

    @classmethod
    def get_instance(cls, dict_path=None):
        if cls._instance is None:
            cls._instance = cls(dict_path)
        return cls._instance

    def __init__(self, dict_path=None):
        self.known_terms = []
        self.direct_replacements = {}
        self.standard_units = []
        self.known_words = set()
        self._load_dictionary(dict_path)

    def _load_dictionary(self, dict_path):
        if dict_path is None:
            dict_path = get_dictionary_path()
        loaded = False
        if dict_path and os.path.isfile(dict_path):
            try:
                with open(dict_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.known_terms = data.get("known_terms", [])
                self.direct_replacements = data.get("direct_replacements", {})
                self.standard_units = data.get("standard_units", [])
                for phrase in self.known_terms:
                    for w in phrase.split():
                        w_clean = re.sub(r"[^\w\u0600-\u06FF]", "", w).strip()
                        if len(w_clean) >= 3:
                            self.known_words.add(w_clean)
                loaded = True
            except Exception as e:
                print(f"Warning: Failed to parse dictionary at {dict_path}: {e}")

        if not loaded:
            self.direct_replacements = {
                "مواسر": "مواسير",
                "مفاس": "مقاس",
                "فطر": "قطر",
                "تطر": "قطر",
                "حالطى": "حائطي",
                "حانطى": "حائطي",
                "بلاستبك": "بلاستيك",
                "نفتيش": "تفتيش",
                "تهدنة": "تهدئة",
                "واختباد": "واختبار",
                "غرنة": "غرفة",
                "سحان": "سخان",
                "مياذكهرى": "مياه كهربائي",
                "راسى": "رأسي"
            }

    def correct_text(self, text: str) -> str:
        if not text:
            return ""

        result = text
        # 1. Direct multi-word phrase replacements
        for src, dst in self.direct_replacements.items():
            if " " in src and src in result:
                result = result.replace(src, dst)

        # 2. Token-level direct replacements & fuzzy matching
        tokens = result.split()
        corrected_tokens = []
        known_words_list = list(self.known_words) if HAS_RAPIDFUZZ else []

        for token in tokens:
            prefix = ""
            suffix = ""
            core = token
            while core and core[0] in "():;,.،؛'\"-":
                prefix += core[0]
                core = core[1:]
            while core and core[-1] in "():;,.،؛'\"-":
                suffix = core[-1] + suffix
                core = core[:-1]

            if not core:
                corrected_tokens.append(token)
                continue

            if core in self.direct_replacements:
                core = self.direct_replacements[core]
            elif core in self.known_words:
                pass
            elif len(core) >= 4 and not any(ch.isdigit() for ch in core) and any("\u0600" <= ch <= "\u06FF" for ch in core):
                if HAS_RAPIDFUZZ and known_words_list:
                    match = process.extractOne(core, known_words_list, scorer=fuzz.ratio, score_cutoff=78)
                    if match:
                        core = match[0]
                else:
                    import difflib
                    matches = difflib.get_close_matches(core, self.known_words, n=1, cutoff=0.85)
                    if matches:
                        core = matches[0]

            corrected_tokens.append(prefix + core + suffix)

        result = " ".join(corrected_tokens)
        result = StructuredTableExtractor.clean_arabic_dimensions(result)
        return result

    def correct_unit(self, unit: str) -> str:
        if not unit:
            return ""
        u = unit.strip().replace(" ", "").replace(".", "")
        if any(k in u for k in ["عدد", "عدذ", "عذد", "عد"]):
            return "عدد"
        if any(k in u for k in ["مط", "م.ط", "ط", "درطط", "م"]):
            return "م.ط"
        if any(k in u for k in ["م2", "متر مربع", "مترمربع"]):
            return "م2"
        if any(k in u for k in ["م3", "مترمكعب", "متر مكعب"]):
            return "م3"
        if any(k in u for k in ["كجم", "كيلو"]):
            return "كجم"
        if any(k in u for k in ["طن"]):
            return "طن"
        if any(k in u for k in ["لتر", "لت"]):
            return "لتر"
        return unit.strip()


# ==============================================================================
# Spatial Table Extraction Engine
# ==============================================================================
class StructuredTableExtractor:
    """
    Extracts structured multi-column table grids from document page images
    and RapidOCR bounding boxes using OpenCV morphology grid detection,
    coordinate row clustering, explicit 3-column boundary locking, and category merging.
    """

    @staticmethod
    def is_arabic(text: str) -> bool:
        return any('\u0600' <= ch <= '\u06FF' for ch in str(text))

    @staticmethod
    def clean_arabic_dimensions(s: str) -> str:
        """Standardizes dimensions, technical acronyms, and numeric sequences."""
        if not s:
            return ""
        s = re.sub(r"^([A-Za-z0-9\-_]{2,})\s+([\u0600-\u06FF]+)", r"\2 \1", s)
        s = re.sub(r"\b([A-Za-z0-9\-_]{2,})\s+(مواسير|محبس|طوب|صرف|غرفة|سخان|كرفان|الواح|زجاج)\b", r"\2 \1", s)
        s = re.sub(r"(\d+)\s*[\'’`\"*xX×*٢،,]\s*(\d+)\s*[\'’`\"*xX×*٢،,]\s*(\d+)\s*(?:س?م|م?م|م)?", r"\1*\2*\3 سم", s)
        s = re.sub(r"(\d+)\s*[\'’`\"*xX×*٢،,]\s*(\d+)\s*(س?م|م?م|م)?\b", r"\1*\2 سم", s)
        s = re.sub(r"(\d+)\s+(\d+)\s*س\b", r"\1*\2 سم", s)
        s = re.sub(r"(\d+)\s+(\d+)\s*سم\b", r"\1*\2 سم", s)
        s = re.sub(r"سمعمق", "سم عمق", s)
        s = re.sub(r"عمق\s*(\d+)\s*(?:س?م)?\b", r"عمق \1 سم", s)
        s = re.sub(r"(\d+)\s*مع\b", r"\1 مم", s)
        s = re.sub(r"(\d+)\s*م~\b", r"\1 مم", s)
        s = re.sub(r"(\d+)\s*م~~\b", r"\1 مم", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    @classmethod
    def assemble_cell_text(cls, items_list, is_rtl=True):
        if not items_list:
            return ""
        if len(items_list) == 1:
            t = items_list[0]['text']
            return cls.clean_arabic_dimensions(t) if is_rtl else t

        items_sorted = sorted(items_list, key=lambda x: x['yc'])
        lines = []
        for it in items_sorted:
            placed = False
            for line in lines:
                line_yc = sum(x['yc'] for x in line) / len(line)
                if abs(it['yc'] - line_yc) < 14:
                    line.append(it)
                    placed = True
                    break
            if not placed:
                lines.append([it])

        line_texts = []
        for line in lines:
            if is_rtl:
                line.sort(key=lambda x: x['xc'], reverse=True)
            else:
                line.sort(key=lambda x: x['xc'])
            line_str = " ".join(x['text'] for x in line)
            if is_rtl:
                line_str = cls.clean_arabic_dimensions(line_str)
            line_texts.append(line_str)

        res = " ".join(line_texts)
        if is_rtl:
            res = cls.clean_arabic_dimensions(res)
        return res

    @classmethod
    def clean_unit(cls, u: str) -> str:
        if not u:
            return ""
        u_clean = u.strip().replace(" ", "").replace(".", "")
        if any(k in u_clean for k in ["عدد", "عدذ", "عذد", "عد"]):
            return "عدد"
        if any(k in u_clean for k in ["مط", "م.ط", "ط", "درطط", "م"]):
            return "م.ط"
        return u.strip()

    @classmethod
    def extract_table_from_page(cls, img_bgr, ocr_results, apply_dictionary=True):
        """
        Extracts structured table rows and columns from OCR bounding boxes using
        OpenCV morphology horizontal lines and strict X-coordinate column boundaries.
        """
        if not ocr_results:
            return {
                'title_headers': [],
                'table_rows': [],
                'category_rows': set(),
                'footer_notes': [],
                'is_rtl': False,
                'has_grid': False
            }

        if img_bgr is not None:
            img_h, img_w = img_bgr.shape[:2]
        else:
            all_xs = [pt[0] for bbox, _, _ in ocr_results for pt in bbox]
            all_ys = [pt[1] for bbox, _, _ in ocr_results for pt in bbox]
            img_w = max(all_xs) + 50 if all_xs else 1653
            img_h = max(all_ys) + 50 if all_ys else 2339

        items = []
        arabic_chars = 0
        total_chars = 0
        for bbox, text, conf in ocr_results:
            t = str(text).strip()
            if not t or t.lower() == "camscanner":
                continue
            xs = [pt[0] for pt in bbox]
            ys = [pt[1] for pt in bbox]
            w = max(xs) - min(xs)
            h = max(ys) - min(ys)
            total_chars += len(t)
            if cls.is_arabic(t):
                arabic_chars += len(t)

            # Filter noise boxes
            if float(conf) < 0.20 and not any(c.isalnum() for c in t):
                continue
            if h > (img_h * 0.025) and w < (img_w * 0.03) and t in (":", ".", "|", "", "-", "=", "ة"):
                continue

            items.append({
                'x0': min(xs), 'x1': max(xs),
                'y0': min(ys), 'y1': max(ys),
                'xc': sum(xs) / len(xs),
                'yc': sum(ys) / len(ys),
                'h': h, 'w': w,
                'text': t, 'conf': float(conf)
            })

        is_rtl = (arabic_chars / max(1, total_chars)) > 0.20

        # Horizontal Grid Line Detection via OpenCV morphology
        line_ys = []
        has_physical_grid = False
        if img_bgr is not None and HAS_CV2:
            try:
                gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (img_w // 10, 1))
                horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
                h_proj = np.sum(horizontal > 0, axis=1)

                in_line = False
                start_y = 0
                for y in range(img_h):
                    if h_proj[y] > (img_w * 0.40):
                        if not in_line:
                            in_line = True
                            start_y = y
                    else:
                        if in_line:
                            in_line = False
                            line_ys.append((start_y + y - 1) / 2)
                if len(line_ys) >= 6:
                    has_physical_grid = True
            except Exception:
                has_physical_grid = False

        title_headers = []
        footer_notes = []
        table_bands = []

        if has_physical_grid:
            first_grid_y = line_ys[0]
            last_grid_y = line_ys[-1]

            top_items = [it for it in items if it['yc'] < first_grid_y]
            if top_items:
                title_headers.append(cls.clean_arabic_dimensions(" ".join(x['text'] for x in sorted(top_items, key=lambda x: x['xc'], reverse=True))))

            bot_items = [it for it in items if it['yc'] > last_grid_y]
            if bot_items:
                footer_notes.append(" ".join(x['text'] for x in bot_items))

            raw_bands = []
            for i in range(len(line_ys) - 1):
                y_t = line_ys[i]
                y_b = line_ys[i + 1]
                if y_b - y_t < 15:
                    continue
                if y_b - y_t > 60:
                    band_items = [it for it in items if y_t <= it['yc'] < y_b]
                    sub_ycs = sorted(list(set(round(it['yc'] / 20) * 20 for it in band_items)))
                    if len(sub_ycs) > 1:
                        mid_y = (y_t + y_b) / 2
                        raw_bands.append((y_t, mid_y))
                        raw_bands.append((mid_y, y_b))
                        continue
                raw_bands.append((y_t, y_b))

            for y_t, y_b in raw_bands:
                r_items = [it for it in items if y_t <= it['yc'] < y_b]
                table_bands.append((r_items, y_t, y_b))
        else:
            # Fallback vertical spatial clustering
            items_sorted = sorted(items, key=lambda it: it['yc'])
            row_clusters = []
            for it in items_sorted:
                matched = None
                for r in row_clusters:
                    r_yc = sum(x['yc'] for x in r) / len(r)
                    r_h = sum(x['h'] for x in r) / len(r)
                    overlap = max(0, min(it['y1'], max(x['y1'] for x in r)) - max(it['y0'], min(x['y0'] for x in r)))
                    min_h = min(it['h'], r_h)
                    if abs(it['yc'] - r_yc) < 10 or (abs(it['yc'] - r_yc) < 14 and min_h > 0 and overlap / min_h > 0.40):
                        matched = r
                        break
                if matched is not None:
                    matched.append(it)
                else:
                    row_clusters.append([it])

            for r in row_clusters:
                y_t = min(x['y0'] for x in r)
                y_b = max(x['y1'] for x in r)
                if y_t < (img_h * 0.058):
                    title_headers.append(cls.assemble_cell_text(r, is_rtl))
                    continue
                if y_b > (img_h * 0.96):
                    footer_notes.append(cls.assemble_cell_text(r, is_rtl))
                    continue
                table_bands.append((r, y_t, y_b))

        # Column coordinate thresholds: Price (left), Unit (center), Description (right)
        x_price_thresh = img_w * 0.24
        x_unit_thresh = img_w * 0.42

        raw_rows = []
        category_indices = set()
        corrector = ArabicDictionaryCorrector.get_instance() if apply_dictionary else None

        for r_items, y_t, y_b in table_bands:
            if not r_items:
                continue

            # Category banner detection: wide box or section header keyword
            is_cat_banner = False
            for it in r_items:
                w_ratio = it['w'] / img_w
                if (w_ratio > 0.35 and any(k in it['text'] for k in ["مصنوعية", "مصنعية", "مواسير", "اعمال", "أعمال", "تمديدات", "تغذية", "عزل"])) or \
                   (it['xc'] > (img_w * 0.30) and any(it['text'].startswith(k) for k in ["مصنوعية", "مصنعية"])):
                    is_cat_banner = True
                    break

            if is_cat_banner:
                banner_text = cls.assemble_cell_text(r_items, is_rtl=True)
                if corrector:
                    banner_text = corrector.correct_text(banner_text)
                category_indices.add(len(raw_rows))
                raw_rows.append([banner_text, "", ""])
                continue

            desc_items = []
            unit_items = []
            price_items = []

            for it in r_items:
                if it['xc'] < x_price_thresh:
                    price_items.append(it)
                elif it['xc'] < x_unit_thresh:
                    unit_items.append(it)
                else:
                    desc_items.append(it)

            desc_text = cls.assemble_cell_text(desc_items, is_rtl)
            unit_text = cls.clean_unit(" ".join(x['text'] for x in sorted(unit_items, key=lambda x: x['xc'])))
            price_text = " ".join(x['text'] for x in sorted(price_items, key=lambda x: x['xc'])).strip()

            if corrector:
                desc_text = corrector.correct_text(desc_text)
                unit_text = corrector.correct_unit(unit_text)

            raw_rows.append([desc_text, unit_text, price_text])

        # Filter out empty phantom rows
        final_rows = []
        final_cat_indices = set()
        for idx, row in enumerate(raw_rows):
            desc, unit, price = row
            if idx in category_indices:
                if desc.strip():
                    final_cat_indices.add(len(final_rows))
                    final_rows.append([desc.strip(), "", ""])
                continue

            if not desc.strip():
                continue
            if len(desc.strip()) <= 1 and not price.strip():
                continue

            final_rows.append([desc.strip(), unit.strip(), price.strip()])

        return {
            'title_headers': title_headers,
            'table_rows': final_rows,
            'category_rows': final_cat_indices,
            'footer_notes': footer_notes,
            'is_rtl': is_rtl,
            'has_grid': has_physical_grid
        }


# ==============================================================================
# Multiprocessing Worker Initialization & Job Execution
# ==============================================================================
_GLOBAL_RAPIDOCR = None
_GLOBAL_CORRECTOR = None


def init_worker_process(models_dir: str = None, dict_path: str = None):
    """Initializes RapidOCR engine and Egyptian Arabic dictionary once per child process."""
    global _GLOBAL_RAPIDOCR, _GLOBAL_CORRECTOR
    try:
        if models_dir is None:
            models_dir = get_model_storage_dir()
        if dict_path is None:
            dict_path = get_dictionary_path()
        _GLOBAL_RAPIDOCR = create_rapidocr_engine(models_dir)
        _GLOBAL_CORRECTOR = ArabicDictionaryCorrector.get_instance(dict_path)
    except Exception as e:
        print(f"[Worker Init] Error: {e}", file=sys.stderr)


def process_pdf_page_job(page_idx: int, pdf_path: str, models_dir: str = None, dict_path: str = None, apply_dictionary: bool = True, dpi: int = 200):
    """
    Picklable multiprocessing job for parsing a single PDF page:
    1. Zero-copy rasterization via PyMuPDF.
    2. High-speed RapidOCR ONNX inference on raw image buffer.
    3. RTL Arabic normalization and spatial table extraction.
    """
    global _GLOBAL_RAPIDOCR, _GLOBAL_CORRECTOR
    if _GLOBAL_RAPIDOCR is None:
        init_worker_process(models_dir, dict_path)

    import pymupdf
    import numpy as np

    doc = pymupdf.open(pdf_path)
    try:
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=dpi)
        raw = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        img_bgr = raw[:, :, :3] if pix.n >= 3 else raw

        raw_ocr, _ = _GLOBAL_RAPIDOCR(img_bgr)
        ocr_results = []
        if raw_ocr:
            for box, text, score in raw_ocr:
                norm_text = normalize_rapidocr_arabic(text)
                if norm_text:
                    ocr_results.append((box, norm_text, float(score)))

        table_data = StructuredTableExtractor.extract_table_from_page(
            img_bgr,
            ocr_results,
            apply_dictionary=apply_dictionary
        )

        block_count = len(ocr_results)
        char_count = sum(len(str(it[1])) for it in ocr_results)
        simple_ocr = [(list(it[0]), str(it[1]), float(it[2])) for it in ocr_results]

        return {
            'page_num': page_idx + 1,
            'page_idx': page_idx,
            'table_data': table_data,
            'ocr_results': simple_ocr,
            'block_count': block_count,
            'char_count': char_count,
            'error': None
        }
    except Exception as e:
        return {
            'page_num': page_idx + 1,
            'page_idx': page_idx,
            'table_data': {'title_headers': [], 'table_rows': [], 'category_rows': set(), 'footer_notes': [], 'is_rtl': True, 'has_grid': False},
            'ocr_results': [],
            'block_count': 0,
            'char_count': 0,
            'error': str(e)
        }
    finally:
        doc.close()


# ==============================================================================
# Conversion Engine & File Exporter
# ==============================================================================
class ConversionEngineMixin:
    """Shared document parsing, OCR pipeline, and multi-column file export helpers."""

    def __init__(self):
        self.ocr_engine = None
        self.last_converted_path = None

    def _get_ocr_engine(self):
        if self.ocr_engine is None:
            self.ocr_engine = create_rapidocr_engine(get_model_storage_dir())
        return self.ocr_engine

    def _save_excel(self, extracted_pages_data, output_path, input_path):
        if not HAS_OPENPYXL:
            raise RuntimeError("openpyxl is not installed. Run 'pip install openpyxl'.")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        title_font = Font(name="Segoe UI", size=13, bold=True, color="1E3A8A")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")

        cat_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        cat_font = Font(name="Segoe UI", size=11, bold=True, color="1E293B")

        alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        data_font = Font(name="Segoe UI", size=10, color="0F172A")

        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        right_align = Alignment(horizontal="right", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        thin_side = Side(style="thin", color="CBD5E1")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        cat_border_right = Border(top=thin_side, bottom=thin_side, right=thin_side)
        cat_border_mid = Border(top=thin_side, bottom=thin_side)
        cat_border_left = Border(top=thin_side, bottom=thin_side, left=thin_side)

        # 1. Consolidated Sheet
        ws_all = wb.create_sheet(title="All Extracted Data")
        if any(p['table_data'].get('is_rtl') for p in extracted_pages_data):
            try:
                ws_all.sheet_view.rightToLeft = True
            except Exception:
                pass

        current_row = 1
        for page_info in extracted_pages_data:
            page_num = page_info['page_num']
            t_data = page_info['table_data']
            is_rtl = t_data.get('is_rtl', True)
            align_text = right_align if is_rtl else left_align
            category_rows = t_data.get('category_rows', set())

            ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
            banner = ws_all.cell(row=current_row, column=1)
            banner.value = f"--- [ صفحة / Page {page_num} ] ---"
            banner.font = Font(name="Segoe UI", size=11, bold=True, color="475569")
            banner.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            banner.alignment = center_align
            ws_all.row_dimensions[current_row].height = 28
            current_row += 1

            for h_text in t_data.get('title_headers', []):
                ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                c = ws_all.cell(row=current_row, column=1)
                c.value = h_text
                c.font = title_font
                c.alignment = center_align
                ws_all.row_dimensions[current_row].height = 26
                current_row += 1

            for r_i, row in enumerate(t_data.get('table_rows', [])):
                is_first_row = (r_i == 0)
                is_category = (r_i in category_rows)

                if is_category:
                    ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                    c1 = ws_all.cell(row=current_row, column=1)
                    c1.value = row[0]
                    c1.font = cat_font
                    c1.fill = cat_fill
                    c1.alignment = center_align
                    c1.border = cat_border_right

                    c2 = ws_all.cell(row=current_row, column=2)
                    c2.fill = cat_fill
                    c2.border = cat_border_mid

                    c3 = ws_all.cell(row=current_row, column=3)
                    c3.fill = cat_fill
                    c3.border = cat_border_left
                    ws_all.row_dimensions[current_row].height = 26
                else:
                    for c_i in range(3):
                        col_idx = c_i + 1
                        cell_val = row[c_i] if c_i < len(row) else ""
                        c = ws_all.cell(row=current_row, column=col_idx)
                        c.value = cell_val
                        c.border = thin_border
                        if is_first_row:
                            c.font = header_font
                            c.fill = header_fill
                            c.alignment = center_align
                        else:
                            c.font = data_font
                            if r_i % 2 == 1:
                                c.fill = alt_fill
                            if c_i in (1, 2):
                                c.alignment = center_align
                            else:
                                c.alignment = align_text
                    ws_all.row_dimensions[current_row].height = 26 if is_first_row else 22
                current_row += 1

            for f_text in t_data.get('footer_notes', []):
                ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                c = ws_all.cell(row=current_row, column=1, value=f_text)
                c.font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
                c.alignment = center_align
                current_row += 1

            current_row += 1

        ws_all.column_dimensions["A"].width = 54
        ws_all.column_dimensions["B"].width = 16
        ws_all.column_dimensions["C"].width = 18

        # 2. Individual Per-Page Sheets
        for page_info in extracted_pages_data:
            page_num = page_info['page_num']
            t_data = page_info['table_data']
            is_rtl = t_data.get('is_rtl', True)
            align_text = right_align if is_rtl else left_align
            category_rows = t_data.get('category_rows', set())

            ws_page = wb.create_sheet(title=f"Page {page_num}")
            if is_rtl:
                try:
                    ws_page.sheet_view.rightToLeft = True
                except Exception:
                    pass

            p_row = 1
            for h_text in t_data.get('title_headers', []):
                ws_page.merge_cells(start_row=p_row, start_column=1, end_row=p_row, end_column=3)
                c = ws_page.cell(row=p_row, column=1, value=h_text)
                c.font = title_font
                c.alignment = center_align
                ws_page.row_dimensions[p_row].height = 26
                p_row += 1

            for r_i, row in enumerate(t_data.get('table_rows', [])):
                is_first_row = (r_i == 0)
                is_category = (r_i in category_rows)

                if is_category:
                    ws_page.merge_cells(start_row=p_row, start_column=1, end_row=p_row, end_column=3)
                    c1 = ws_page.cell(row=p_row, column=1, value=row[0])
                    c1.font = cat_font
                    c1.fill = cat_fill
                    c1.alignment = center_align
                    c1.border = cat_border_right

                    c2 = ws_page.cell(row=p_row, column=2)
                    c2.fill = cat_fill
                    c2.border = cat_border_mid

                    c3 = ws_page.cell(row=p_row, column=3)
                    c3.fill = cat_fill
                    c3.border = cat_border_left
                    ws_page.row_dimensions[p_row].height = 26
                else:
                    for c_i in range(3):
                        col_idx = c_i + 1
                        cell_val = row[c_i] if c_i < len(row) else ""
                        c = ws_page.cell(row=p_row, column=col_idx, value=cell_val)
                        c.border = thin_border
                        if is_first_row:
                            c.font = header_font
                            c.fill = header_fill
                            c.alignment = center_align
                        else:
                            c.font = data_font
                            if r_i % 2 == 1:
                                c.fill = alt_fill
                            if c_i in (1, 2):
                                c.alignment = center_align
                            else:
                                c.alignment = align_text
                    ws_page.row_dimensions[p_row].height = 26 if is_first_row else 22
                p_row += 1

            for f_text in t_data.get('footer_notes', []):
                ws_page.merge_cells(start_row=p_row, start_column=1, end_row=p_row, end_column=3)
                c = ws_page.cell(row=p_row, column=1, value=f_text)
                c.font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
                c.alignment = center_align
                p_row += 1

            ws_page.column_dimensions["A"].width = 54
            ws_page.column_dimensions["B"].width = 16
            ws_page.column_dimensions["C"].width = 18

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        wb.save(output_path)

    def _save_word(self, extracted_pages_data, output_path, input_path):
        if not HAS_DOCX:
            raise RuntimeError("python-docx is not installed. Run 'pip install python-docx'.")

        doc = docx.Document()
        doc.add_heading(os.path.basename(input_path), level=0)
        total_pages = len(extracted_pages_data)
        sub = doc.add_paragraph(f"Extracted with PDF2text • {total_pages} page(s)")
        if sub.runs:
            sub.runs[0].font.italic = True
            sub.runs[0].font.color.rgb = RGBColor(100, 116, 139)

        for page_info in extracted_pages_data:
            page_num = page_info['page_num']
            t_data = page_info['table_data']
            doc.add_heading(f"Page {page_num}", level=1)

            for h in t_data.get('title_headers', []):
                p = doc.add_paragraph()
                r = p.add_run(h)
                r.bold = True

            rows = t_data.get('table_rows', [])
            if rows:
                table = doc.add_table(rows=len(rows), cols=3)
                table.style = 'Light Shading Accent 1'
                category_rows = t_data.get('category_rows', set())

                for r_idx, row in enumerate(rows):
                    row_cells = table.rows[r_idx].cells
                    if r_idx in category_rows:
                        row_cells[0].merge(row_cells[2])
                        row_cells[0].text = row[0]
                    else:
                        for c_idx in range(min(3, len(row))):
                            row_cells[c_idx].text = str(row[c_idx])

            for f in t_data.get('footer_notes', []):
                p = doc.add_paragraph()
                r = p.add_run(f)
                r.italic = True

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        doc.save(output_path)

    def _save_text(self, extracted_pages_data, output_path, input_path):
        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"=== {os.path.basename(input_path)} ===\n\n")
            for page_info in extracted_pages_data:
                page_num = page_info['page_num']
                t_data = page_info['table_data']
                f.write(f"--- [ صفحة / Page {page_num} ] ---\n")
                for h in t_data.get('title_headers', []):
                    f.write(f"[Header] {h}\n")
                for row in t_data.get('table_rows', []):
                    f.write("\t".join(str(cell) for cell in row) + "\n")
                for note in t_data.get('footer_notes', []):
                    f.write(f"[Footer] {note}\n")
                f.write("\n")


# ==============================================================================
# Interactive Page Preview Dialog
# ==============================================================================
class PagePreviewDialog(tk.Toplevel):
    """
    Modal window displaying interactive preview of Page 1 table extraction.
    Allows editing cells, toggling section banners, and reviewing layout before full export.
    """

    def __init__(self, parent, page_1_info, total_pages, target_format, on_decision):
        super().__init__(parent)
        self.page_1_info = page_1_info
        self.total_pages = total_pages
        self.target_format = target_format
        self.on_decision = on_decision
        self.decision_made = False

        self.title("PDF2text - Page 1 Extraction Preview")
        self.geometry("860x580")
        self.minsize(700, 480)
        self.configure(bg="#F8FAFC")
        self.transient(parent)
        self.grab_set()

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._create_widgets()
        self._populate_table()
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _create_widgets(self):
        top_frame = tk.Frame(self, bg="#FFFFFF", padx=20, pady=12, highlightthickness=1, highlightbackground="#E2E8F0")
        top_frame.pack(fill="x")

        title_lbl = tk.Label(
            top_frame,
            text="Interactive Page 1 Table Preview",
            font=("Segoe UI", 12, "bold"),
            fg="#1E293B",
            bg="#FFFFFF"
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            top_frame,
            text=f"Review extracted columns. Remaining {self.total_pages - 1} page(s) are prefetching concurrently in the background.",
            font=("Segoe UI", 9),
            fg="#64748B",
            bg="#FFFFFF"
        )
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # Table container
        tree_frame = tk.Frame(self, bg="#F8FAFC", padx=16, pady=10)
        tree_frame.pack(fill="both", expand=True)

        cols = ("idx", "desc", "unit", "price", "type")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")

        self.tree.heading("idx", text="#", anchor="center")
        self.tree.heading("desc", text="Description / البند", anchor="e")
        self.tree.heading("unit", text="Unit / الوحدة", anchor="center")
        self.tree.heading("price", text="Price / السعر", anchor="center")
        self.tree.heading("type", text="Row Type", anchor="center")

        self.tree.column("idx", width=45, anchor="center", stretch=False)
        self.tree.column("desc", width=440, anchor="e")
        self.tree.column("unit", width=95, anchor="center", stretch=False)
        self.tree.column("price", width=105, anchor="center", stretch=False)
        self.tree.column("type", width=115, anchor="center", stretch=False)

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.tag_configure("even_row", background="#FFFFFF")
        self.tree.tag_configure("odd_row", background="#F8FAFC")
        self.tree.tag_configure("header_row", background="#1E3A8A", foreground="#FFFFFF")
        self.tree.tag_configure("section_row", background="#FEF08A", foreground="#1E293B")

        self.tree.bind("<Double-1>", self._on_double_click)

        # Toolbar
        bar = tk.Frame(self, bg="#F8FAFC", padx=16, pady=8)
        bar.pack(fill="x")

        self.btn_edit = tk.Button(bar, text="Edit Row", font=("Segoe UI", 9), relief="solid", bd=1, padx=10, pady=4, command=self._edit_selected_row)
        self.btn_edit.pack(side="left", padx=(0, 6))

        self.btn_cat = tk.Button(bar, text="Toggle Banner", font=("Segoe UI", 9), relief="solid", bd=1, padx=10, pady=4, command=self._toggle_category)
        self.btn_cat.pack(side="left", padx=(0, 6))

        self.btn_del = tk.Button(bar, text="Delete Row", font=("Segoe UI", 9), relief="solid", bd=1, padx=10, pady=4, command=self._delete_row)
        self.btn_del.pack(side="left", padx=(0, 6))

        # Action buttons
        btn_box = tk.Frame(self, bg="#FFFFFF", padx=16, pady=12, highlightthickness=1, highlightbackground="#E2E8F0")
        btn_box.pack(fill="x", side="bottom")

        btn_cancel = tk.Button(btn_box, text="Cancel", font=("Segoe UI", 9), relief="flat", padx=16, pady=6, bg="#E2E8F0", fg="#475569", command=self._on_cancel)
        btn_cancel.pack(side="left")

        btn_accept = tk.Button(btn_box, text=f"Confirm & Export ({self.target_format})", font=("Segoe UI", 9, "bold"), relief="flat", padx=18, pady=6, bg="#2563EB", fg="#FFFFFF", command=self._on_accept)
        btn_accept.pack(side="right")

    def _populate_table(self):
        t_data = self.page_1_info.get('table_data', {})
        rows = t_data.get('table_rows', [])
        cat_rows = t_data.get('category_rows', set())

        for idx, row in enumerate(rows):
            desc = row[0] if len(row) > 0 else ""
            unit = row[1] if len(row) > 1 else ""
            price = row[2] if len(row) > 2 else ""

            is_header = (idx == 0)
            is_cat = (idx in cat_rows)

            if is_header:
                tag = "header_row"
                row_type = "Header"
            elif is_cat:
                tag = "section_row"
                row_type = "Category Banner"
                unit = "[Merged]"
                price = "[Merged]"
            else:
                tag = "even_row" if idx % 2 == 0 else "odd_row"
                row_type = "Data"

            self.tree.insert("", "end", iid=str(idx), values=(str(idx + 1), desc, unit, price, row_type), tags=(tag,))

    def _on_double_click(self, event):
        self._edit_selected_row()

    def _edit_selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = list(self.tree.item(item, "values"))

        dlg = tk.Toplevel(self)
        dlg.title("Edit Row")
        dlg.geometry("440x260")
        dlg.configure(bg="#FFFFFF", padx=16, pady=14)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Description / البند:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF").pack(anchor="w")
        e_desc = tk.Entry(dlg, font=("Segoe UI", 10), justify="right")
        e_desc.insert(0, vals[1])
        e_desc.pack(fill="x", pady=(2, 8))

        tk.Label(dlg, text="Unit / الوحدة:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF").pack(anchor="w")
        e_unit = tk.Entry(dlg, font=("Segoe UI", 10), justify="center")
        e_unit.insert(0, "" if vals[2] == "[Merged]" else vals[2])
        e_unit.pack(fill="x", pady=(2, 8))

        tk.Label(dlg, text="Price / السعر:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF").pack(anchor="w")
        e_price = tk.Entry(dlg, font=("Segoe UI", 10), justify="center")
        e_price.insert(0, "" if vals[3] == "[Merged]" else vals[3])
        e_price.pack(fill="x", pady=(2, 8))

        is_cat_var = tk.BooleanVar(value=(vals[4] == "Category Banner"))
        chk_cat = tk.Checkbutton(dlg, text="Is Full-Width Category Banner", variable=is_cat_var, bg="#FFFFFF", font=("Segoe UI", 9))
        chk_cat.pack(anchor="w", pady=(0, 10))

        def _save():
            vals[1] = e_desc.get().strip()
            if is_cat_var.get():
                vals[2] = "[Merged]"
                vals[3] = "[Merged]"
                vals[4] = "Category Banner"
                self.tree.item(item, values=vals, tags=("section_row",))
            else:
                vals[2] = e_unit.get().strip()
                vals[3] = e_price.get().strip()
                vals[4] = "Data"
                self.tree.item(item, values=vals, tags=("even_row",))
            dlg.destroy()

        btn_save = tk.Button(dlg, text="Save", font=("Segoe UI", 9, "bold"), bg="#2563EB", fg="#FFFFFF", relief="flat", padx=14, pady=5, command=_save)
        btn_save.pack(side="right")

    def _toggle_category(self):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = list(self.tree.item(item, "values"))
        if vals[4] == "Category Banner":
            vals[4] = "Data"
            vals[2] = "عدد"
            vals[3] = ""
            self.tree.item(item, values=vals, tags=("even_row",))
        else:
            vals[4] = "Category Banner"
            vals[2] = "[Merged]"
            vals[3] = "[Merged]"
            self.tree.item(item, values=vals, tags=("section_row",))

    def _delete_row(self):
        sel = self.tree.selection()
        if sel:
            self.tree.delete(sel[0])

    def _on_accept(self):
        if not self.decision_made:
            self.decision_made = True
            rows = []
            cats = set()
            for idx, child in enumerate(self.tree.get_children()):
                v = self.tree.item(child, "values")
                is_cat = (v[4] == "Category Banner")
                if is_cat:
                    cats.add(idx)
                    rows.append([v[1], "", ""])
                else:
                    rows.append([v[1], v[2] if v[2] != "[Merged]" else "", v[3] if v[3] != "[Merged]" else ""])

            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
            if self.on_decision:
                self.on_decision({'accepted': True, 'table_rows': rows, 'category_rows': cats})

    def _on_cancel(self):
        if not self.decision_made:
            self.decision_made = True
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
            if self.on_decision:
                self.on_decision({'accepted': False})


# ==============================================================================
# Modern Desktop GUI (CustomTkinter with Tkinter Fallback)
# ==============================================================================
_BaseAppClass = ctk.CTk if HAS_CTK else tk.Tk

class PDF2textApp(_BaseAppClass, ConversionEngineMixin):
    """Modern Desktop GUI Application for PDF2text Document & OCR Converter."""

    def __init__(self):
        _BaseAppClass.__init__(self)
        ConversionEngineMixin.__init__(self)

        if HAS_CTK:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")
            self.configure(fg_color="#F1F5F9")

        self.title("PDF2text - Desktop Document & OCR Converter")
        self.geometry("620x670")
        self.minsize(560, 600)

        self.input_file_path = None
        self.output_file_path = None
        self.selected_format = "Excel (.xlsx)"
        self.is_processing = False

        self._setup_styles()
        self._create_widgets()
        self._center_window()

    def _setup_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _create_widgets(self):
        # Card container
        if HAS_CTK:
            self.card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=16, border_width=1, border_color="#E2E8F0")
            self.card.pack(fill="both", expand=True, padx=24, pady=24)
        else:
            self.card = tk.Frame(self, bg="#FFFFFF", padx=24, pady=24)
            self.card.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header_frame = tk.Frame(self.card, bg="#FFFFFF")
        header_frame.pack(fill="x", pady=(0, 16))

        title_lbl = tk.Label(header_frame, text="PDF2text", font=("Segoe UI", 20, "bold"), fg="#1E293B", bg="#FFFFFF")
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header_frame,
            text="High-Speed Arabic & English Document & OCR Converter (ONNX Runtime)",
            font=("Segoe UI", 10),
            fg="#64748B",
            bg="#FFFFFF"
        )
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # Input File Section
        in_lbl = tk.Label(self.card, text="Input Document (PDF, Image, or Scanned File):", font=("Segoe UI", 10, "bold"), fg="#334155", bg="#FFFFFF")
        in_lbl.pack(anchor="w", pady=(4, 4))

        in_box = tk.Frame(self.card, bg="#FFFFFF")
        in_box.pack(fill="x", pady=(0, 14))

        self.txt_input = tk.Entry(in_box, font=("Segoe UI", 10), bg="#F8FAFC", relief="solid", bd=1, fg="#0F172A")
        self.txt_input.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        btn_browse_in = tk.Button(in_box, text="Browse File", font=("Segoe UI", 9, "bold"), bg="#2563EB", fg="#FFFFFF", relief="flat", padx=14, pady=6, command=self._on_browse_input)
        btn_browse_in.pack(side="right")

        # Format Section
        fmt_lbl = tk.Label(self.card, text="Target Export Format:", font=("Segoe UI", 10, "bold"), fg="#334155", bg="#FFFFFF")
        fmt_lbl.pack(anchor="w", pady=(4, 4))

        self.fmt_var = tk.StringVar(value=self.selected_format)
        self.cmb_format = ttk.Combobox(self.card, textvariable=self.fmt_var, values=FORMAT_OPTIONS, state="readonly", font=("Segoe UI", 10))
        self.cmb_format.pack(fill="x", pady=(0, 14), ipady=3)
        self.cmb_format.bind("<<ComboboxSelected>>", self._on_format_changed)

        # Output File Section
        out_lbl = tk.Label(self.card, text="Destination File Path:", font=("Segoe UI", 10, "bold"), fg="#334155", bg="#FFFFFF")
        out_lbl.pack(anchor="w", pady=(4, 4))

        out_box = tk.Frame(self.card, bg="#FFFFFF")
        out_box.pack(fill="x", pady=(0, 16))

        self.txt_output = tk.Entry(out_box, font=("Segoe UI", 10), bg="#F8FAFC", relief="solid", bd=1, fg="#0F172A")
        self.txt_output.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        btn_browse_out = tk.Button(out_box, text="Change Save Location", font=("Segoe UI", 9), bg="#E2E8F0", fg="#334155", relief="flat", padx=12, pady=6, command=self._on_browse_output)
        btn_browse_out.pack(side="right")

        # Convert Action Button
        self.btn_convert = tk.Button(
            self.card,
            text="Convert Document Now",
            font=("Segoe UI", 11, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            relief="flat",
            pady=10,
            command=self._on_start_conversion
        )
        self.btn_convert.pack(fill="x", pady=(6, 16))

        # Progress bar
        self.prog_bar = ttk.Progressbar(self.card, orient="horizontal", mode="determinate")
        self.prog_bar.pack(fill="x", pady=(0, 8))

        # Status text
        self.lbl_status = tk.Label(self.card, text="Ready to convert documents.", font=("Segoe UI", 9, "bold"), fg="#475569", bg="#FFFFFF", anchor="w")
        self.lbl_status.pack(fill="x")

        self.lbl_detail = tk.Label(self.card, text="Select an input PDF or image file and click 'Convert Document Now'.", font=("Segoe UI", 8), fg="#94A3B8", bg="#FFFFFF", anchor="w")
        self.lbl_detail.pack(fill="x", pady=(2, 10))

        # Success Action Box
        self.action_box = tk.Frame(self.card, bg="#F0FDF4", padx=12, pady=10, highlightthickness=1, highlightbackground="#BBF7D0")
        self.btn_open_file = tk.Button(self.action_box, text="Open File", font=("Segoe UI", 9, "bold"), bg="#16A34A", fg="#FFFFFF", relief="flat", padx=14, pady=5, command=self._open_converted_file)
        self.btn_open_file.pack(side="left", padx=(0, 8))

        self.btn_open_folder = tk.Button(self.action_box, text="Open Folder", font=("Segoe UI", 9), bg="#DCFCE7", fg="#166534", relief="flat", padx=12, pady=5, command=self._open_output_folder)
        self.btn_open_folder.pack(side="left")

    def _on_browse_input(self):
        f = filedialog.askopenfilename(
            title="Select Document to Convert",
            filetypes=[
                ("All Supported Documents", "*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff"),
                ("PDF Documents (*.pdf)", "*.pdf"),
                ("Images (*.png, *.jpg, *.jpeg)", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff"),
                ("All Files (*.*)", "*.*")
            ]
        )
        if f:
            self.input_file_path = f
            self.txt_input.delete(0, "end")
            self.txt_input.insert(0, f)

            # Auto-suggest output path
            ext = FORMAT_CONFIG[self.fmt_var.get()]["ext"]
            out_path = os.path.splitext(f)[0] + ext
            self.output_file_path = out_path
            self.txt_output.delete(0, "end")
            self.txt_output.insert(0, out_path)

            self.action_box.pack_forget()
            self.lbl_status.configure(text=f"Loaded: {os.path.basename(f)}", fg="#2563EB")
            self.lbl_detail.configure(text="Ready to begin conversion. Click 'Convert Document Now'.", fg="#64748B")

    def _on_format_changed(self, event=None):
        self.selected_format = self.fmt_var.get()
        if self.txt_input.get():
            in_p = self.txt_input.get()
            ext = FORMAT_CONFIG[self.selected_format]["ext"]
            out_p = os.path.splitext(in_p)[0] + ext
            self.txt_output.delete(0, "end")
            self.txt_output.insert(0, out_p)
            self.output_file_path = out_p

    def _on_browse_output(self):
        cfg = FORMAT_CONFIG.get(self.fmt_var.get(), FORMAT_CONFIG["Excel (.xlsx)"])
        f = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=cfg["ext"],
            filetypes=cfg["filetypes"]
        )
        if f:
            self.output_file_path = f
            self.txt_output.delete(0, "end")
            self.txt_output.insert(0, f)

    def _update_status_safe(self, message=None, progress=None, detail=None, text_color=None):
        def _apply():
            if message is not None:
                self.lbl_status.configure(text=message)
            if detail is not None:
                self.lbl_detail.configure(text=detail)
            if progress is not None:
                self.prog_bar["value"] = progress * 100
            if text_color is not None:
                self.lbl_status.configure(fg=text_color)
        self.after(0, _apply)

    def _on_start_conversion(self):
        input_path = self.txt_input.get().strip()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input document.")
            return

        target_format = self.fmt_var.get()
        target_ext = FORMAT_CONFIG[target_format]["ext"]
        output_path = self.txt_output.get().strip()
        if not output_path:
            output_path = os.path.splitext(input_path)[0] + target_ext
            self.txt_output.delete(0, "end")
            self.txt_output.insert(0, output_path)

        if os.path.exists(output_path):
            confirm = messagebox.askyesno(
                "Confirm Overwrite",
                f"The destination file already exists:\n\n{os.path.basename(output_path)}\n\nDo you want to overwrite it?",
                icon="warning"
            )
            if not confirm:
                return

        self.btn_convert.configure(state="disabled", text="Processing Document...")
        self.action_box.pack_forget()

        thread = threading.Thread(
            target=self._run_conversion_worker,
            args=(input_path, output_path, target_format),
            daemon=True
        )
        thread.start()

    def _open_page_1_preview(self, page_1_info, total_pages, target_format, on_decision):
        try:
            PagePreviewDialog(self, page_1_info, total_pages, target_format, on_decision)
        except Exception as e:
            print(f"Preview Dialog Exception: {e}")
            on_decision(True)

    def _on_conversion_cancelled(self):
        self.btn_convert.configure(state="normal", text="Convert Document Now")
        self._update_status_safe(
            message="Operation cancelled by user.",
            progress=0.0,
            detail="No files were exported. Ready for next document.",
            text_color="#64748B"
        )

    def _run_conversion_worker(self, input_path: str, output_path: str, target_format: str):
        start_time = time.time()
        try:
            self._update_status_safe(
                message="[1/4] Initializing RapidOCR (ONNX Runtime)...",
                progress=0.08,
                detail="Loading high-speed Arabic ONNX recognizer...",
                text_color="#2563EB"
            )
            engine = self._get_ocr_engine()

            ext = os.path.splitext(input_path)[1].lower()
            self._update_status_safe(
                message=f"[2/4] Opening document: {os.path.basename(input_path)}...",
                progress=0.15,
                detail="Zero-copy rasterization via PyMuPDF...",
                text_color="#2563EB"
            )

            extracted_pages_data = []

            if ext == ".pdf":
                if not HAS_PYMUPDF:
                    raise RuntimeError("PyMuPDF is required to process PDF documents.")

                doc = pymupdf.open(input_path)
                total_pages = len(doc)

                # Process Page 1 immediately
                self._update_status_safe(
                    message=f"[3/4] Parsing Page 1 of {total_pages}...",
                    progress=0.20,
                    detail="Running RapidOCR on Page 1 image...",
                    text_color="#2563EB"
                )
                page0 = doc[0]
                pix0 = page0.get_pixmap(dpi=200)
                raw0 = np.frombuffer(pix0.samples, dtype=np.uint8).reshape(pix0.h, pix0.w, pix0.n)
                img0 = raw0[:, :, :3] if pix0.n >= 3 else raw0

                raw_ocr0, _ = engine(img0)
                ocr_res0 = []
                if raw_ocr0:
                    for box, text, score in raw_ocr0:
                        norm_text = normalize_rapidocr_arabic(text)
                        if norm_text:
                            ocr_res0.append((box, norm_text, float(score)))

                table_data0 = StructuredTableExtractor.extract_table_from_page(img0, ocr_res0, apply_dictionary=True)
                page_1_info = {
                    'page_num': 1,
                    'table_data': table_data0,
                    'ocr_results': ocr_res0
                }
                extracted_pages_data.append(page_1_info)

                # Launch Background Prefetch Pipeline for Pages 2..N
                pool_executor = None
                prefetch_futures = {}
                prefetch_results = {}
                prefetch_cancelled = threading.Event()
                num_workers = min(4, max(1, os.cpu_count() // 2))

                if total_pages > 1:
                    models_dir = get_model_storage_dir()
                    dict_path = get_dictionary_path()
                    try:
                        pool_executor = ProcessPoolExecutor(
                            max_workers=num_workers,
                            initializer=init_worker_process,
                            initargs=(models_dir, dict_path)
                        )
                        for p_idx in range(1, total_pages):
                            fut = pool_executor.submit(
                                process_pdf_page_job,
                                p_idx,
                                input_path,
                                models_dir,
                                dict_path,
                                True,
                                200
                            )
                            prefetch_futures[fut] = p_idx

                        def _monitor_prefetch():
                            completed_count = 0
                            for f in as_completed(prefetch_futures):
                                if prefetch_cancelled.is_set():
                                    break
                                p_idx = prefetch_futures[f]
                                try:
                                    res = f.result()
                                    prefetch_results[p_idx] = res
                                    completed_count += 1
                                    self._update_status_safe(
                                        detail=f"Previewing Page 1 • Prefetched {completed_count}/{total_pages - 1} background pages...",
                                        text_color="#059669"
                                    )
                                except Exception as e:
                                    print(f"Prefetch page {p_idx+1} error: {e}")

                        threading.Thread(target=_monitor_prefetch, daemon=True).start()
                    except Exception as pool_err:
                        print(f"ProcessPoolExecutor fallback: {pool_err}")
                        pool_executor = None

                # Open interactive preview modal
                preview_event = threading.Event()
                preview_result = {'accepted': False}

                def _on_preview_choice(choice):
                    if isinstance(choice, dict):
                        preview_result.update(choice)
                    else:
                        preview_result['accepted'] = bool(choice)
                    preview_event.set()

                self._update_status_safe(
                    message=f"[Preview] Page 1 ready ({len(table_data0['table_rows'])} rows). Review while remaining pages prefetch...",
                    progress=0.30,
                    detail=f"Multi-core prefetch active across {num_workers} CPU cores. Review Page 1 table layout.",
                    text_color="#D97706"
                )
                self.after(0, self._open_page_1_preview, page_1_info, total_pages, target_format, _on_preview_choice)

                preview_event.wait()

                if not preview_result.get('accepted', False):
                    prefetch_cancelled.set()
                    if pool_executor:
                        pool_executor.shutdown(wait=False, cancel_futures=True)
                    doc.close()
                    self.after(0, self._on_conversion_cancelled)
                    return

                if 'table_rows' in preview_result and preview_result['table_rows']:
                    page_1_info['table_data']['table_rows'] = preview_result['table_rows']
                if 'category_rows' in preview_result and preview_result['category_rows'] is not None:
                    page_1_info['table_data']['category_rows'] = preview_result['category_rows']

                # Finalize remaining pages
                self._update_status_safe(
                    message=f"[3/4] Finalizing remaining pages...",
                    progress=0.85,
                    detail="Assembling multi-core prefetched pages...",
                    text_color="#2563EB"
                )

                if pool_executor and prefetch_futures:
                    for fut, p_idx in prefetch_futures.items():
                        if p_idx not in prefetch_results:
                            try:
                                res = fut.result()
                                prefetch_results[p_idx] = res
                            except Exception as e:
                                print(f"Page {p_idx+1} execution error: {e}")
                    pool_executor.shutdown(wait=True)

                    for p_idx in range(1, total_pages):
                        if p_idx in prefetch_results:
                            extracted_pages_data.append(prefetch_results[p_idx])
                elif total_pages > 1:
                    # Sequential fallback
                    for p_idx in range(1, total_pages):
                        page_res = process_pdf_page_job(p_idx, input_path, get_model_storage_dir(), get_dictionary_path())
                        extracted_pages_data.append(page_res)

                doc.close()
            else:
                # Single Image File
                img_bgr = cv2.imread(input_path) if HAS_CV2 else None
                raw_ocr, _ = engine(img_bgr)
                ocr_res = []
                if raw_ocr:
                    for box, text, score in raw_ocr:
                        norm_text = normalize_rapidocr_arabic(text)
                        if norm_text:
                            ocr_res.append((box, norm_text, float(score)))

                t_data = StructuredTableExtractor.extract_table_from_page(img_bgr, ocr_res, apply_dictionary=True)
                extracted_pages_data.append({
                    'page_num': 1,
                    'table_data': t_data,
                    'ocr_results': ocr_res
                })

            # Step 4: Export to Target Format
            self._update_status_safe(
                message=f"[4/4] Writing output file: {os.path.basename(output_path)}...",
                progress=0.92,
                detail=f"Formatting {target_format} layout...",
                text_color="#2563EB"
            )

            if target_format == "Excel (.xlsx)":
                self._save_excel(extracted_pages_data, output_path, input_path)
            elif target_format == "Word (.docx)":
                self._save_word(extracted_pages_data, output_path, input_path)
            elif target_format == "Text (.txt)":
                self._save_text(extracted_pages_data, output_path, input_path)

            elapsed = time.time() - start_time
            self.last_converted_path = output_path

            def _on_success():
                self.btn_convert.configure(state="normal", text="Convert Document Now")
                self.action_box.pack(fill="x", pady=(0, 10))
                total_rows = sum(len(p['table_data'].get('table_rows', [])) for p in extracted_pages_data)
                self.lbl_status.configure(
                    text=f"Conversion Complete! ({elapsed:.1f}s)",
                    fg="#16A34A"
                )
                self.lbl_detail.configure(
                    text=f"Successfully extracted {len(extracted_pages_data)} page(s) and {total_rows} table rows to {os.path.basename(output_path)}.",
                    fg="#059669"
                )
                self.prog_bar["value"] = 100

            self.after(0, _on_success)

        except Exception as e:
            elapsed = time.time() - start_time
            print(f"Conversion Worker Exception: {e}", file=sys.stderr)

            def _on_error():
                self.btn_convert.configure(state="normal", text="Convert Document Now")
                self.lbl_status.configure(
                    text="Conversion failed. An error occurred.",
                    fg="#DC2626"
                )
                self.lbl_detail.configure(
                    text=str(e),
                    fg="#EF4444"
                )
                messagebox.showerror("Conversion Error", f"Failed to convert document:\n\n{str(e)}")

            self.after(0, _on_error)

    def _open_converted_file(self):
        if self.last_converted_path and os.path.exists(self.last_converted_path):
            try:
                os.startfile(self.last_converted_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {e}")

    def _open_output_folder(self):
        if self.last_converted_path and os.path.exists(self.last_converted_path):
            folder = os.path.dirname(os.path.abspath(self.last_converted_path))
            try:
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open folder: {e}")


def main():
    """Main entry point with Windows multiprocessing freeze support."""
    multiprocessing.freeze_support()
    app = PDF2textApp()
    app.mainloop()


if __name__ == "__main__":
    main()
