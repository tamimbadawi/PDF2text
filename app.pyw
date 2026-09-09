"""
PDF2text - Desktop Document & OCR Converter
Converts PDF, scanned documents, and images to Excel (.xlsx), Word (.docx), and Text (.txt).
Features high-speed RapidOCR (ONNX Runtime), PyMuPDF zero-copy rendering,
spatial table extraction with category banner merging, and Egyptian construction dictionary.
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

# Modern CustomTkinter UI styling
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

FILE_TYPES = [
    ("All Supported Documents", "*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff"),
    ("PDF Documents (*.pdf)", "*.pdf"),
    ("Images (*.png, *.jpg, *.jpeg)", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff"),
    ("All Files (*.*)", "*.*")
]


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
    - Reversing the string restores natural Arabic reading order.
    - Re-inverts numbers, dimensions, and Latin symbols to maintain natural LTR order.
    """
    if not raw_text:
        return ""
    t = raw_text.strip()
    if not any("\u0600" <= ch <= "\u06FF" for ch in t):
        return t

    rev = t[::-1]
    fixed = re.sub(r"([A-Za-z0-9\-_./*×+]+)", lambda m: m.group(1)[::-1], rev)
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
        for src, dst in self.direct_replacements.items():
            if " " in src and src in result:
                result = result.replace(src, dst)

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
    and RapidOCR bounding boxes using vertical proximity row clustering,
    explicit 3-column coordinate boundary locking, and section category banner merging.
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
        Extracts structured table rows, columns, and section category banners from OCR bounding boxes.
        Uses vertical proximity clustering and column boundary thresholds.
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

        # Form precise row clusters by vertical proximity
        items.sort(key=lambda x: x['yc'])
        clusters = []
        for it in items:
            placed = False
            for cl in clusters:
                avg_yc = sum(x['yc'] for x in cl) / len(cl)
                if abs(it['yc'] - avg_yc) < 13:
                    cl.append(it)
                    placed = True
                    break
            if not placed:
                clusters.append([it])

        # Column coordinate thresholds: Price (left), Unit (center), Description (right)
        x_price_thresh = img_w * 0.24
        x_unit_thresh = img_w * 0.42

        extracted_rows = []
        category_row_indices = set()
        title_headers = []
        footer_notes = []

        corrector = ArabicDictionaryCorrector.get_instance() if apply_dictionary else None

        for cl in clusters:
            avg_yc = sum(x['yc'] for x in cl) / len(cl)
            row_text = " ".join(x['text'] for x in sorted(cl, key=lambda x: x['xc'], reverse=True))

            # Header above the table
            if any(k in row_text for k in ["البند", "وحدة القياس", "متوسط السعر"]) and avg_yc < 200:
                extracted_rows.append(["البند", "وحدة القياس", "متوسط السعر"])
                continue

            price_items = [x for x in cl if x['xc'] < x_price_thresh]
            unit_items = [x for x in cl if x_price_thresh <= x['xc'] < x_unit_thresh]
            desc_items = [x for x in cl if x['xc'] >= x_unit_thresh]

            # Check for section category banner
            is_cat = False
            row_xc = sum(x['xc'] for x in cl) / len(cl)

            if any(k in row_text for k in ["مصنعية الغرف", "مصنوعية الغرف", "السخانات", "تغذية", "حائطى بول", "حائطي بول", "صرف حائط", "UPVC", "سيملس"]):
                if not any(k in row_text for k in ["قطر خارج", "قطر 25", "قطر 32", "قطر 50", "بوصة"]):
                    is_cat = True
            elif not price_items and not unit_items:
                if (img_w * 0.28 <= row_xc <= img_w * 0.72):
                    if not any(k in row_text for k in ["قطر", "مقاس", "مفاس", "بوصة"]):
                        is_cat = True

            if is_cat:
                cat_title = corrector.correct_text(row_text) if corrector else row_text
                category_row_indices.add(len(extracted_rows))
                extracted_rows.append([cat_title, "", ""])
                continue

            desc_str = " ".join(x['text'] for x in sorted(desc_items, key=lambda x: x['xc'], reverse=True))
            unit_str = cls.clean_unit(" ".join(x['text'] for x in unit_items))
            price_str = " ".join(x['text'] for x in sorted(price_items, key=lambda x: x['xc'])).strip()

            if corrector:
                desc_str = corrector.correct_text(desc_str)
                unit_str = corrector.correct_unit(unit_str)

            if not desc_str and not price_str:
                continue
            if len(desc_str) <= 1 and not price_str:
                continue

            extracted_rows.append([desc_str, unit_str, price_str])

        # Contextual Unit Inheritance across section items
        active_unit = "عدد"
        for idx, row in enumerate(extracted_rows):
            if idx in category_row_indices:
                desc_lower = row[0]
                if any(k in desc_lower for k in ["مواسير", "صرف", "UPVC", "سيملس", "HDPE", "بولي"]):
                    active_unit = "م.ط"
                else:
                    active_unit = "عدد"
            elif idx > 0:  # Skip table header
                if not row[1]:
                    row[1] = active_unit
                else:
                    active_unit = row[1]

        return {
            'title_headers': title_headers,
            'table_rows': extracted_rows,
            'category_rows': category_row_indices,
            'footer_notes': footer_notes,
            'is_rtl': is_rtl,
            'has_grid': True
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
                c = ws_all.cell(row=current_row, column=1, value=h_text)
                c.font = title_font
                c.alignment = center_align
                ws_all.row_dimensions[current_row].height = 26
                current_row += 1

            for r_i, row in enumerate(t_data.get('table_rows', [])):
                is_first_row = (r_i == 0)
                is_category = (r_i in category_rows)

                if is_category:
                    ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                    c1 = ws_all.cell(row=current_row, column=1, value=row[0])
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
                        c = ws_all.cell(row=current_row, column=col_idx, value=cell_val)
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


class PagePreviewDialog(tk.Toplevel):
    """
    Interactive modal dialog presenting an editable structured table preview of Page 1.
    Allows user to inspect and edit column mapping (Description, Unit, Price) and section headers.
    Provides options to apply dictionary spell-correction to remaining pages.
    """
    def __init__(self, parent, page_data, total_pages, target_format, on_decision):
        super().__init__(parent)
        self.on_decision = on_decision
        self.total_pages = total_pages
        self.target_format = target_format
        self.page_data = page_data
        self.decision_made = False

        self.apply_dict_var = tk.BooleanVar(value=True)
        self.edit_entry = None
        self.edit_item = None
        self.edit_col = None

        self.title("PDF2text - Page 1 Structured Preview & Interactive Correction")
        self.geometry("980x680")
        self.minsize(800, 540)
        self.configure(bg="#F8FAFC")

        # Modal configuration
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass

        self._setup_ui()
        self._populate_data()
        self._center_window(parent)

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self, parent):
        self.update_idletasks()
        w = 980
        h = 680
        try:
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            x = max(0, px + (pw - w) // 2)
            y = max(0, py + (ph - h) // 2)
        except Exception:
            x = (self.winfo_screenwidth() - w) // 2
            y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _setup_ui(self):
        # Header banner
        header_frame = tk.Frame(self, bg="#FFFFFF", padx=24, pady=16, highlightthickness=1, highlightbackground="#E2E8F0")
        header_frame.pack(fill="x")

        title_lbl = tk.Label(
            header_frame,
            text="Page 1 Interactive Preview & Spell Correction",
            font=("Segoe UI", 16, "bold"),
            bg="#FFFFFF",
            fg="#0F172A"
        )
        title_lbl.pack(anchor="w")

        sub_text = (
            f"Review extracted table grid before processing remaining pages. "
            f"Total Document Pages: {self.total_pages} • Target Format: {self.target_format}"
        )
        sub_lbl = tk.Label(
            header_frame,
            text=sub_text,
            font=("Segoe UI", 10),
            bg="#FFFFFF",
            fg="#64748B"
        )
        sub_lbl.pack(anchor="w", pady=(2, 8))

        t_data = self.page_data.get('table_data', {})
        t_rows = t_data.get('table_rows', [])
        cat_rows = t_data.get('category_rows', set())
        data_count = max(0, len(t_rows) - len(cat_rows) - 1)

        pills_frame = tk.Frame(header_frame, bg="#FFFFFF")
        pills_frame.pack(anchor="w")

        pills = [
            (f"Total Rows: {len(t_rows)}", "#EFF6FF", "#1E40AF"),
            (f"Section Categories: {len(cat_rows)}", "#FEF3C7", "#92400E"),
            (f"Data Items: {data_count}", "#F0FDF4", "#166534"),
            ("3 Columns: Description • Unit • Price", "#F1F5F9", "#334155"),
            ("✓ Egyptian Dictionary Active", "#ECFDF5", "#047857")
        ]
        for text, bg_c, fg_c in pills:
            lbl = tk.Label(
                pills_frame,
                text=text,
                font=("Segoe UI", 9, "bold"),
                bg=bg_c,
                fg=fg_c,
                padx=10,
                pady=3
            )
            lbl.pack(side="left", padx=(0, 8))

        # Quick Actions & Options Toolbar
        tb_frame = tk.Frame(self, bg="#F8FAFC", padx=20, pady=8)
        tb_frame.pack(fill="x")

        btn_edit = tk.Button(
            tb_frame,
            text="✏ Edit Selected Cell/Row",
            font=("Segoe UI", 9, "bold"),
            bg="#FFFFFF",
            fg="#1E293B",
            relief="solid",
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._edit_selected_row
        )
        btn_edit.pack(side="left", padx=(0, 6))

        btn_toggle = tk.Button(
            tb_frame,
            text="🔄 Toggle Category Banner",
            font=("Segoe UI", 9, "bold"),
            bg="#FEF08A",
            fg="#854D0E",
            relief="solid",
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._toggle_category
        )
        btn_toggle.pack(side="left", padx=(0, 6))

        btn_del = tk.Button(
            tb_frame,
            text="🗑 Delete Row",
            font=("Segoe UI", 9),
            bg="#FEE2E2",
            fg="#991B1B",
            relief="solid",
            bd=1,
            padx=8,
            pady=4,
            cursor="hand2",
            command=self._delete_row
        )
        btn_del.pack(side="left", padx=(0, 6))

        # Dictionary option checkbox
        chk = tk.Checkbutton(
            tb_frame,
            text="Apply Egyptian dictionary spell-corrections to remaining pages",
            variable=self.apply_dict_var,
            font=("Segoe UI", 9, "bold"),
            bg="#F8FAFC",
            fg="#1E3A8A",
            activebackground="#F8FAFC"
        )
        chk.pack(side="right")

        # Body Table Container
        body_frame = tk.Frame(self, bg="#F8FAFC", padx=20, pady=4)
        body_frame.pack(fill="both", expand=True)

        # Style Treeview
        tree_style = ttk.Style(self)
        try:
            tree_style.theme_use("clam")
        except Exception:
            pass
        tree_style.configure(
            "Preview.Treeview",
            font=("Segoe UI", 10),
            rowheight=26,
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground="#0F172A",
            bordercolor="#CBD5E1"
        )
        tree_style.configure(
            "Preview.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#F1F5F9",
            foreground="#0F172A",
            padding=(6, 6)
        )
        tree_style.map("Preview.Treeview.Heading", background=[("active", "#E2E8F0")])

        # Table Treeview
        cols = ("row", "desc", "unit", "price", "type")
        self.tree = ttk.Treeview(body_frame, columns=cols, show="headings", style="Preview.Treeview", selectmode="browse")

        self.tree.heading("row", text="#", anchor="center")
        self.tree.heading("desc", text="Description / البند (Double-click to edit)", anchor="e")
        self.tree.heading("unit", text="Unit / وحدة القياس", anchor="center")
        self.tree.heading("price", text="Price / متوسط السعر", anchor="center")
        self.tree.heading("type", text="Layout Type", anchor="center")

        self.tree.column("row", width=45, minwidth=40, anchor="center")
        self.tree.column("desc", width=480, minwidth=300, anchor="e")
        self.tree.column("unit", width=120, minwidth=90, anchor="center")
        self.tree.column("price", width=110, minwidth=90, anchor="center")
        self.tree.column("type", width=130, minwidth=100, anchor="center")

        # Scrollbars
        vsb = ttk.Scrollbar(body_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(body_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        body_frame.grid_rowconfigure(0, weight=1)
        body_frame.grid_columnconfigure(0, weight=1)

        # Tags styling
        self.tree.tag_configure("header_row", background="#1E3A8A", foreground="#FFFFFF", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("section_row", background="#FEF08A", foreground="#1E293B", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("even_row", background="#FFFFFF", foreground="#0F172A")
        self.tree.tag_configure("odd_row", background="#F8FAFC", foreground="#0F172A")

        # Double click to edit cell inline
        self.tree.bind("<Double-1>", self._on_double_click)

        # Bottom Actions Bar
        actions_bar = tk.Frame(self, bg="#FFFFFF", padx=24, pady=14, highlightthickness=1, highlightbackground="#E2E8F0")
        actions_bar.pack(fill="x", side="bottom")

        hint_lbl = tk.Label(
            actions_bar,
            text="💡 Tip: Double-click any cell or press 'Edit Selected' to adjust values or numbers.",
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#64748B"
        )
        hint_lbl.pack(side="left")

        btn_box = tk.Frame(actions_bar, bg="#FFFFFF")
        btn_box.pack(side="right")

        # Cancel button
        cancel_btn = tk.Button(
            btn_box,
            text="✖ Cancel",
            font=("Segoe UI", 10, "bold"),
            bg="#F1F5F9",
            fg="#1E293B",
            activebackground="#E2E8F0",
            activeforeground="#0F172A",
            relief="flat",
            bd=0,
            padx=16,
            pady=8,
            cursor="hand2",
            command=self._on_cancel
        )
        cancel_btn.pack(side="left", padx=(0, 10))

        # Accept button
        accept_text = f"✔ Accept & Process Remaining Pages ({self.total_pages - 1})" if self.total_pages > 1 else "✔ Accept & Export Document"
        accept_btn = tk.Button(
            btn_box,
            text=accept_text,
            font=("Segoe UI", 10, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._on_accept
        )
        accept_btn.pack(side="left")

    def _populate_data(self):
        t_data = self.page_data.get('table_data', {})
        t_rows = t_data.get('table_rows', [])
        cat_rows = t_data.get('category_rows', set())

        for idx, row in enumerate(t_rows):
            desc = row[0] if len(row) > 0 else ""
            unit = row[1] if len(row) > 1 else ""
            price = row[2] if len(row) > 2 else ""

            is_first = (idx == 0)
            is_cat = (idx in cat_rows)

            if is_first:
                tag = "header_row"
                row_type = "Table Header"
            elif is_cat:
                tag = "section_row"
                row_type = "Section Category"
                unit = "[Merged Across]"
                price = "[Merged Across]"
            else:
                tag = "odd_row" if idx % 2 == 1 else "even_row"
                row_type = "Data Row"

            self.tree.insert("", "end", values=(idx + 1, desc, unit, price, row_type), tags=(tag,))

    def _on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item or not column:
            return

        col_idx = int(column.replace("#", "")) - 1
        if col_idx == 0:  # Row number cannot be edited
            return

        self._spawn_inline_editor(item, col_idx)

    def _spawn_inline_editor(self, item, col_idx):
        if self.edit_entry:
            self._save_inline_edit()

        bbox = self.tree.bbox(item, f"#{col_idx + 1}")
        if not bbox:
            return

        curr_val = self.tree.item(item, "values")[col_idx]
        if curr_val == "[Merged Across]":
            curr_val = ""

        entry = tk.Entry(self.tree, font=("Segoe UI", 10), justify="center" if col_idx in (2, 3) else "right")
        entry.insert(0, curr_val)
        entry.select_range(0, "end")
        entry.focus_set()

        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

        self.edit_entry = entry
        self.edit_item = item
        self.edit_col = col_idx

        entry.bind("<Return>", lambda e: self._save_inline_edit())
        entry.bind("<FocusOut>", lambda e: self._save_inline_edit())
        entry.bind("<Escape>", lambda e: self._cancel_inline_edit())

    def _save_inline_edit(self):
        if not self.edit_entry or not self.edit_item:
            return
        new_val = self.edit_entry.get().strip()
        vals = list(self.tree.item(self.edit_item, "values"))
        vals[self.edit_col] = new_val

        # If user edited layout type
        if self.edit_col == 4:
            if "Category" in new_val or "قسم" in new_val:
                vals[2] = "[Merged Across]"
                vals[3] = "[Merged Across]"
                vals[4] = "Section Category"
                self.tree.item(self.edit_item, tags=("section_row",))
            else:
                vals[4] = "Data Row"
                self.tree.item(self.edit_item, tags=("even_row",))

        self.tree.item(self.edit_item, values=vals)
        self.edit_entry.destroy()
        self.edit_entry = None
        self.edit_item = None
        self.edit_col = None

    def _cancel_inline_edit(self):
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
            self.edit_item = None
            self.edit_col = None

    def _edit_selected_row(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Select Row", "Please select a row in the table first.", parent=self)
            return
        item = selected[0]
        vals = list(self.tree.item(item, "values"))

        dlg = tk.Toplevel(self)
        dlg.title("Edit Row Data")
        dlg.geometry("450x270")
        dlg.configure(bg="#FFFFFF", padx=20, pady=16)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Description / البند:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF").pack(anchor="w")
        e_desc = tk.Entry(dlg, font=("Segoe UI", 10), justify="right")
        e_desc.insert(0, vals[1])
        e_desc.pack(fill="x", pady=(2, 10))

        tk.Label(dlg, text="Unit / وحدة القياس:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF").pack(anchor="w")
        e_unit = tk.Entry(dlg, font=("Segoe UI", 10), justify="center")
        e_unit.insert(0, "" if vals[2] == "[Merged Across]" else vals[2])
        e_unit.pack(fill="x", pady=(2, 10))

        tk.Label(dlg, text="Price / متوسط السعر:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF").pack(anchor="w")
        e_price = tk.Entry(dlg, font=("Segoe UI", 10), justify="center")
        e_price.insert(0, "" if vals[3] == "[Merged Across]" else vals[3])
        e_price.pack(fill="x", pady=(2, 10))

        is_cat_var = tk.BooleanVar(value=(vals[4] == "Section Category"))
        chk_cat = tk.Checkbutton(dlg, text="Is Full-Width Section Category (Merged Banner)", variable=is_cat_var, bg="#FFFFFF", font=("Segoe UI", 9))
        chk_cat.pack(anchor="w", pady=(0, 12))

        def _save():
            vals[1] = e_desc.get().strip()
            if is_cat_var.get():
                vals[2] = "[Merged Across]"
                vals[3] = "[Merged Across]"
                vals[4] = "Section Category"
                self.tree.item(item, values=vals, tags=("section_row",))
            else:
                vals[2] = e_unit.get().strip()
                vals[3] = e_price.get().strip()
                vals[4] = "Data Row"
                self.tree.item(item, values=vals, tags=("even_row",))
            dlg.destroy()

        btn_save = tk.Button(dlg, text="Save Changes", font=("Segoe UI", 9, "bold"), bg="#2563EB", fg="#FFFFFF", relief="flat", padx=14, pady=6, command=_save)
        btn_save.pack(side="right")

    def _toggle_category(self):
        selected = self.tree.selection()
        if not selected:
            return
        item = selected[0]
        vals = list(self.tree.item(item, "values"))
        if vals[4] == "Section Category":
            vals[4] = "Data Row"
            vals[2] = "م.ط"
            vals[3] = ""
            self.tree.item(item, values=vals, tags=("even_row",))
        else:
            vals[4] = "Section Category"
            vals[2] = "[Merged Across]"
            vals[3] = "[Merged Across]"
            self.tree.item(item, values=vals, tags=("section_row",))

    def _delete_row(self):
        selected = self.tree.selection()
        if not selected:
            return
        self.tree.delete(selected[0])

    def _on_accept(self):
        if not self.decision_made:
            self.decision_made = True
            edited_rows = []
            category_rows = set()
            for idx, item_id in enumerate(self.tree.get_children()):
                vals = self.tree.item(item_id, "values")
                desc = vals[1]
                unit = vals[2] if vals[2] != "[Merged Across]" else ""
                price = vals[3] if vals[3] != "[Merged Across]" else ""
                is_cat = (vals[4] == "Section Category")
                if is_cat:
                    category_rows.add(idx)
                edited_rows.append([desc, unit, price])

            result = {
                'accepted': True,
                'table_rows': edited_rows,
                'category_rows': category_rows,
                'apply_dictionary': self.apply_dict_var.get()
            }
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
            if self.on_decision:
                self.on_decision(result)

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



if HAS_CTK:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    class PDF2textApp(ctk.CTk, ConversionEngineMixin):
        def __init__(self):
            ctk.CTk.__init__(self)
            ConversionEngineMixin.__init__(self)

            # Window configuration
            self.title("PDF2text - Document & OCR Converter")
            self.geometry("620x670")
            self.minsize(560, 600)
            self.configure(fg_color="#F1F5F9")  # Soft modern slate background

            self._setup_styles()
            self._create_widgets()
            self._center_window()

        def _setup_styles(self):
            self.style = ttk.Style()
            try:
                self.style.theme_use("clam")
            except Exception:
                pass

            self.style.configure(
                "Format.TCombobox",
                font=("Segoe UI", 11),
                background="#FFFFFF",
                fieldbackground="#FFFFFF",
                foreground="#0F172A",
                bordercolor="#CBD5E1",
                lightcolor="#CBD5E1",
                darkcolor="#CBD5E1",
                arrowcolor="#2563EB",
                padding=(8, 6)
            )
            self.style.map(
                "Format.TCombobox",
                fieldbackground=[("readonly", "#FFFFFF"), ("focus", "#FFFFFF")],
                selectbackground=[("readonly", "#2563EB"), ("focus", "#2563EB")],
                selectforeground=[("readonly", "#FFFFFF"), ("focus", "#FFFFFF")]
            )

        def _center_window(self):
            self.update_idletasks()
            width = self.winfo_width()
            height = self.winfo_height()
            x = (self.winfo_screenwidth() // 2) - (width // 2)
            y = (self.winfo_screenheight() // 2) - (height // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")

        def _create_widgets(self):
            # Main elevated white card container
            self.card = ctk.CTkFrame(
                self,
                fg_color="#FFFFFF",
                corner_radius=16,
                border_width=1,
                border_color="#E2E8F0"
            )
            self.card.pack(padx=24, pady=20, fill="both", expand=True)

            # --- Header Section ---
            self.header_frame = ctk.CTkFrame(self.card, fg_color="transparent")
            self.header_frame.pack(fill="x", padx=28, pady=(20, 14))

            self.title_label = ctk.CTkLabel(
                self.header_frame,
                text="PDF2text",
                font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
                text_color="#0F172A"
            )
            self.title_label.pack(anchor="w")

            self.subtitle_label = ctk.CTkLabel(
                self.header_frame,
                text="Convert scanned PDF documents and images to structured Excel, Word, or Text",
                font=ctk.CTkFont(family="Segoe UI", size=13),
                text_color="#64748B"
            )
            self.subtitle_label.pack(anchor="w", pady=(2, 0))

            # Subtle Divider
            self.divider = ctk.CTkFrame(self.card, height=1, fg_color="#F1F5F9")
            self.divider.pack(fill="x", padx=28, pady=(0, 16))

            # Form Body
            self.body_frame = ctk.CTkFrame(self.card, fg_color="transparent")
            self.body_frame.pack(fill="both", expand=True, padx=28)

            # --- 1. Input File Section ---
            self.input_label = ctk.CTkLabel(
                self.body_frame,
                text="Input File",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color="#1E293B"
            )
            self.input_label.pack(anchor="w", pady=(0, 4))

            self.input_row = ctk.CTkFrame(self.body_frame, fg_color="transparent")
            self.input_row.pack(fill="x", pady=(0, 14))

            self.input_entry = ctk.CTkEntry(
                self.input_row,
                placeholder_text="Select a PDF or document (e.g. C:/Documents/report.pdf)",
                height=38,
                corner_radius=8,
                border_width=1,
                border_color="#CBD5E1",
                fg_color="#FFFFFF",
                text_color="#0F172A",
                font=ctk.CTkFont(family="Segoe UI", size=12)
            )
            self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

            self.input_browse_btn = ctk.CTkButton(
                self.input_row,
                text="Browse",
                width=90,
                height=38,
                corner_radius=8,
                fg_color="#F1F5F9",
                hover_color="#E2E8F0",
                text_color="#1E293B",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                command=self.browse_input_file
            )
            self.input_browse_btn.pack(side="right")

            # --- 2. Output Format Selector Section (using ttk.Combobox) ---
            self.format_label = ctk.CTkLabel(
                self.body_frame,
                text="Output Format",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color="#1E293B"
            )
            self.format_label.pack(anchor="w", pady=(0, 4))

            self.format_row = ctk.CTkFrame(self.body_frame, fg_color="transparent")
            self.format_row.pack(fill="x", pady=(0, 4))

            self.format_combobox = ttk.Combobox(
                self.format_row,
                values=FORMAT_OPTIONS,
                style="Format.TCombobox",
                state="readonly"
            )
            self.format_combobox.set("Excel (.xlsx)")
            self.format_combobox.pack(side="left", fill="x", expand=True)
            self.format_combobox.bind("<<ComboboxSelected>>", self.on_format_changed)

            self.format_desc_label = ctk.CTkLabel(
                self.body_frame,
                text=FORMAT_CONFIG["Excel (.xlsx)"]["desc"],
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color="#64748B"
            )
            self.format_desc_label.pack(anchor="w", pady=(2, 14))

            # --- 3. Output Destination Section ---
            self.output_label = ctk.CTkLabel(
                self.body_frame,
                text="Output Destination",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color="#1E293B"
            )
            self.output_label.pack(anchor="w", pady=(0, 4))

            self.output_row = ctk.CTkFrame(self.body_frame, fg_color="transparent")
            self.output_row.pack(fill="x", pady=(0, 18))

            self.output_entry = ctk.CTkEntry(
                self.output_row,
                placeholder_text="Destination path (e.g. C:/Documents/report.xlsx)",
                height=38,
                corner_radius=8,
                border_width=1,
                border_color="#CBD5E1",
                fg_color="#FFFFFF",
                text_color="#0F172A",
                font=ctk.CTkFont(family="Segoe UI", size=12)
            )
            self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

            self.output_browse_btn = ctk.CTkButton(
                self.output_row,
                text="Browse",
                width=90,
                height=38,
                corner_radius=8,
                fg_color="#F1F5F9",
                hover_color="#E2E8F0",
                text_color="#1E293B",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                command=self.browse_output_file
            )
            self.output_browse_btn.pack(side="right")

            # --- 4. Convert Button ---
            self.convert_btn = ctk.CTkButton(
                self.body_frame,
                text="Run OCR & Export to Excel (.xlsx)",
                height=44,
                corner_radius=10,
                fg_color="#2563EB",
                hover_color="#1D4ED8",
                text_color="#FFFFFF",
                font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                command=self.start_conversion
            )
            self.convert_btn.pack(fill="x", pady=(0, 12))

            # --- 5. Progress Indicator ---
            self.progress_bar = ctk.CTkProgressBar(
                self.body_frame,
                height=6,
                corner_radius=3,
                progress_color="#2563EB",
                fg_color="#E2E8F0"
            )
            self.progress_bar.pack(fill="x", pady=(0, 8))
            self.progress_bar.set(0)

            # --- 6. Live Status & Dynamic Step Tracking ---
            self.status_label = ctk.CTkLabel(
                self.body_frame,
                text="Ready. Select a document and click Convert.",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color="#334155"
            )
            self.status_label.pack(anchor="w")

            self.detail_label = ctk.CTkLabel(
                self.body_frame,
                text="Step: Idle • Engine: EasyOCR & OpenCV Morphology Table Extraction",
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color="#64748B"
            )
            self.detail_label.pack(anchor="w", pady=(2, 0))

            # --- 7. Post-Conversion Quick Actions ---
            self.actions_row = ctk.CTkFrame(self.body_frame, fg_color="transparent")
            self.actions_row.pack(fill="x", pady=(12, 0))

            self.open_file_btn = ctk.CTkButton(
                self.actions_row,
                text="📄 Open File",
                width=115,
                height=32,
                corner_radius=6,
                fg_color="#F1F5F9",
                hover_color="#E2E8F0",
                text_color="#1E293B",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=self.open_converted_file
            )
            self.open_folder_btn = ctk.CTkButton(
                self.actions_row,
                text="📁 Open Folder",
                width=115,
                height=32,
                corner_radius=6,
                fg_color="#F1F5F9",
                hover_color="#E2E8F0",
                text_color="#1E293B",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=self.open_containing_folder
            )
            self.copy_btn = ctk.CTkButton(
                self.actions_row,
                text="📋 Copy Text",
                width=115,
                height=32,
                corner_radius=6,
                fg_color="#F1F5F9",
                hover_color="#E2E8F0",
                text_color="#1E293B",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=self.copy_to_clipboard
            )

        # ----------------------------------------------------------------------
        # Field & State Accessors
        # ----------------------------------------------------------------------
        def get_input_path(self):
            return self.input_entry.get().strip()

        def set_input_path(self, path):
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, path)

        def get_output_path(self):
            return self.output_entry.get().strip()

        def set_output_path(self, path):
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

        def set_ui_state_processing(self, is_processing: bool):
            if is_processing:
                self.convert_btn.configure(state="disabled", text="Processing...", fg_color="#93C5FD")
                self.input_browse_btn.configure(state="disabled")
                self.output_browse_btn.configure(state="disabled")
                self.format_combobox.configure(state="disabled")
            else:
                fmt = self.format_combobox.get().strip()
                self.convert_btn.configure(state="normal", text=f"Run OCR & Export to {fmt}", fg_color="#2563EB")
                self.input_browse_btn.configure(state="normal")
                self.output_browse_btn.configure(state="normal")
                self.format_combobox.configure(state="readonly")

        def hide_action_buttons(self):
            self.open_file_btn.pack_forget()
            self.open_folder_btn.pack_forget()
            self.copy_btn.pack_forget()

        def show_action_buttons(self):
            self.open_file_btn.pack(side="left", padx=(0, 8))
            self.open_folder_btn.pack(side="left", padx=(0, 8))
            self.copy_btn.pack(side="left")

        # ----------------------------------------------------------------------
        # Live Thread-Safe Status Updates
        # ----------------------------------------------------------------------
        def update_status_safe(self, message: str, progress: float = None, detail: str = None, text_color: str = None):
            """Safely dispatches UI updates onto the Tkinter main thread and renders idletasks."""
            def _apply():
                if message:
                    self.status_label.configure(text=message)
                if text_color:
                    self.status_label.configure(text_color=text_color)
                if detail:
                    self.detail_label.configure(text=detail)
                if progress is not None:
                    clamped = max(0.0, min(1.0, float(progress)))
                    self.progress_bar.set(clamped)
                self.update_idletasks()

            self.after(0, _apply)

        # ----------------------------------------------------------------------
        # Event Handlers & Browse Helpers
        # ----------------------------------------------------------------------
        def on_format_changed(self, event=None):
            fmt = self.format_combobox.get().strip()
            config = FORMAT_CONFIG.get(fmt, FORMAT_CONFIG["Text (.txt)"])
            target_ext = config["ext"]

            # Update convert button text & description hint
            self.convert_btn.configure(text=f"Run OCR & Export to {fmt}")
            self.format_desc_label.configure(text=config["desc"])

            # Synchronize extension in output entry if path is currently present
            curr_out = self.get_output_path()
            if curr_out:
                base, _ = os.path.splitext(curr_out)
                self.set_output_path(base + target_ext)

        def browse_input_file(self):
            selected = filedialog.askopenfilename(
                title="PDF2text - Select Input Document",
                filetypes=FILE_TYPES
            )
            if selected:
                clean_path = os.path.normpath(selected)
                self.set_input_path(clean_path)

                # Automatically suggest output path with the active format extension
                fmt = self.format_combobox.get().strip()
                config = FORMAT_CONFIG.get(fmt, FORMAT_CONFIG["Text (.txt)"])
                target_ext = config["ext"]
                suggested_out = os.path.splitext(clean_path)[0] + target_ext
                self.set_output_path(suggested_out)

                self.update_status_safe(
                    message=f"Document loaded. Ready to convert to {fmt}.",
                    progress=0.0,
                    detail=f"Source: {os.path.basename(clean_path)}",
                    text_color="#334155"
                )

        def browse_output_file(self):
            fmt = self.format_combobox.get().strip()
            config = FORMAT_CONFIG.get(fmt, FORMAT_CONFIG["Text (.txt)"])
            target_ext = config["ext"]

            in_val = self.get_input_path()
            if in_val:
                stem = Path(in_val).stem
                default_name = f"{stem}{target_ext}"
                init_dir = os.path.dirname(os.path.abspath(in_val))
            else:
                default_name = f"output{target_ext}"
                init_dir = os.getcwd()

            selected = filedialog.asksaveasfilename(
                title=f"PDF2text - Save {config['name']}",
                defaultextension=target_ext,
                initialfile=default_name,
                initialdir=init_dir,
                filetypes=config["filetypes"]
            )
            if selected:
                clean_path = os.path.normpath(selected)
                if not clean_path.lower().endswith(target_ext.lower()):
                    clean_path = os.path.splitext(clean_path)[0] + target_ext
                self.set_output_path(clean_path)

        # ----------------------------------------------------------------------
        # Conversion Execution Flow
        # ----------------------------------------------------------------------
        def start_conversion(self):
            input_path = self.get_input_path()
            output_path = self.get_output_path()
            target_format = self.format_combobox.get().strip()
            config = FORMAT_CONFIG.get(target_format, FORMAT_CONFIG["Text (.txt)"])
            target_ext = config["ext"]

            if not input_path:
                messagebox.showwarning("PDF2text", "Please choose an input file to convert.")
                return

            if not os.path.isfile(input_path):
                messagebox.showerror("PDF2text", f"The input file was not found:\n{input_path}")
                return

            if not output_path:
                output_path = os.path.splitext(input_path)[0] + target_ext
                self.set_output_path(output_path)
            elif not output_path.lower().endswith(target_ext.lower()):
                output_path = os.path.splitext(output_path)[0] + target_ext
                self.set_output_path(output_path)

            if os.path.exists(output_path):
                filename = os.path.basename(output_path)
                confirm = messagebox.askyesno(
                    "PDF2text - Confirm Overwrite",
                    f"The destination file already exists:\n\n{filename}\n\nDo you want to overwrite this file?",
                    icon="warning"
                )
                if not confirm:
                    return

            self.set_ui_state_processing(True)
            self.hide_action_buttons()

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
                print(f"Error displaying preview dialog: {e}")
                on_decision(True)

        def _on_conversion_cancelled(self):
            self.set_ui_state_processing(False)

        def _run_conversion_worker(self, input_path: str, output_path: str, target_format: str):
            start_time = time.time()
            try:
                self.update_status_safe(
                    message="[1/4] Initializing RapidOCR (ONNX Runtime)...",
                    progress=0.08,
                    detail="Loading high-speed Arabic ONNX recognizer...",
                    text_color="#2563EB"
                )
                engine = self._get_ocr_engine()

                ext = os.path.splitext(input_path)[1].lower()
                self.update_status_safe(
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
                    self.update_status_safe(
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

                    # Launch Background Prefetch Pipeline for Pages 2..N
                    pool_executor = None
                    prefetch_futures = {}
                    prefetch_results = {}
                    prefetch_cancelled = threading.Event()
                    num_workers = min(4, max(1, (os.cpu_count() or 2) // 2))

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
                                        self.update_status_safe(
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

                    self.update_status_safe(
                        message=f"[Preview] Page 1 ready ({len(table_data0['table_rows'])} rows). Review layout...",
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

                    extracted_pages_data = [page_1_info]

                    # Finalize remaining pages
                    self.update_status_safe(
                        message="[3/4] Finalizing remaining pages...",
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
                        for p_idx in range(1, total_pages):
                            page_res = process_pdf_page_job(p_idx, input_path, get_model_storage_dir(), get_dictionary_path())
                            extracted_pages_data.append(page_res)

                    doc.close()
                else:
                    # Single image file conversion
                    img = cv2.imread(input_path) if HAS_CV2 else None
                    if img is None:
                        raise RuntimeError(f"Could not load image: {input_path}")
                    raw_ocr, _ = engine(img)
                    ocr_res = []
                    if raw_ocr:
                        for box, text, score in raw_ocr:
                            norm_text = normalize_rapidocr_arabic(text)
                            if norm_text:
                                ocr_res.append((box, norm_text, float(score)))
                    table_data = StructuredTableExtractor.extract_table_from_page(img, ocr_res, apply_dictionary=True)
                    page_1_info = {
                        'page_num': 1,
                        'table_data': table_data,
                        'ocr_results': ocr_res
                    }

                    preview_event = threading.Event()
                    preview_result = {'accepted': False}

                    def _on_preview_choice(choice):
                        if isinstance(choice, dict):
                            preview_result.update(choice)
                        else:
                            preview_result['accepted'] = bool(choice)
                        preview_event.set()

                    self.update_status_safe(
                        message=f"[Preview] Image ready ({len(table_data['table_rows'])} rows). Review layout...",
                        progress=0.40,
                        detail="Review detected table structure and columns.",
                        text_color="#D97706"
                    )
                    self.after(0, self._open_page_1_preview, page_1_info, 1, target_format, _on_preview_choice)
                    preview_event.wait()

                    if not preview_result.get('accepted', False):
                        self.after(0, self._on_conversion_cancelled)
                        return

                    if 'table_rows' in preview_result and preview_result['table_rows']:
                        page_1_info['table_data']['table_rows'] = preview_result['table_rows']
                    if 'category_rows' in preview_result and preview_result['category_rows'] is not None:
                        page_1_info['table_data']['category_rows'] = preview_result['category_rows']

                    extracted_pages_data = [page_1_info]

                # Step 4: Export to Target Format
                self.update_status_safe(
                    message=f"[4/4] Exporting to {target_format}...",
                    progress=0.92,
                    detail="Writing structured formatted document...",
                    text_color="#2563EB"
                )

                if "excel" in target_format.lower() or target_format.endswith(".xlsx)"):
                    self._save_excel(extracted_pages_data, output_path, input_path)
                elif "word" in target_format.lower() or target_format.endswith(".docx)"):
                    self._save_word(extracted_pages_data, output_path, input_path)
                else:
                    self._save_text(extracted_pages_data, output_path, input_path)

                self.last_converted_path = output_path
                elapsed = time.time() - start_time
                total_pages_count = len(extracted_pages_data)
                total_rows = sum(len(p['table_data'].get('table_rows', [])) for p in extracted_pages_data)

                self.update_status_safe(
                    message=f"Conversion Complete! ({elapsed:.1f}s)",
                    progress=1.0,
                    detail=f"Saved {total_pages_count} page(s) ({total_rows} rows) to: {os.path.basename(output_path)}",
                    text_color="#16A34A"
                )
                self.after(0, self.show_action_buttons)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.update_status_safe(
                    message="Error during conversion",
                    progress=0.0,
                    detail=f"Details: {str(e)}",
                    text_color="#DC2626"
                )
                messagebox.showerror("PDF2text - Conversion Error", f"An error occurred during conversion:\n\n{str(e)}")
            finally:
                self.after(0, lambda: self.set_ui_state_processing(False))

        def open_converted_file(self):
            if self.last_converted_path and os.path.exists(self.last_converted_path):
                try:
                    os.startfile(self.last_converted_path)
                except Exception as e:
                    messagebox.showerror("PDF2text", f"Could not open file: {e}")

        def open_containing_folder(self):
            if self.last_converted_path and os.path.exists(self.last_converted_path):
                try:
                    folder = os.path.dirname(os.path.abspath(self.last_converted_path))
                    os.startfile(folder)
                except Exception as e:
                    messagebox.showerror("PDF2text", f"Could not open folder: {e}")

        def copy_to_clipboard(self):
            if self.last_converted_path and os.path.exists(self.last_converted_path):
                try:
                    ext = os.path.splitext(self.last_converted_path)[1].lower()
                    if ext == ".txt":
                        with open(self.last_converted_path, "r", encoding="utf-8") as f:
                            text = f.read()
                    else:
                        text = os.path.abspath(self.last_converted_path)

                    self.clipboard_clear()
                    self.clipboard_append(text)
                    self.update_status_safe(detail="Content / file path copied to clipboard!")
                except Exception as e:
                    messagebox.showerror("PDF2text", f"Could not copy text: {e}")


else:
    # Standard Tkinter Fallback for environments without CustomTkinter
    class PDF2textApp(tk.Tk, ConversionEngineMixin):
        def __init__(self):
            super().__init__()
            ConversionEngineMixin.__init__(self)
            self.title("PDF2text - Document & OCR Converter")
            self.geometry("620x670")
            self.minsize(560, 600)
            self.configure(bg="#F1F5F9")
            self.card = tk.Frame(self, bg="#FFFFFF", padx=24, pady=20, highlightthickness=1, highlightbackground="#CBD5E1")
            self.card.pack(padx=20, pady=20, fill="both", expand=True)

            self.lbl_title = tk.Label(self.card, text="PDF2text", font=("Segoe UI", 20, "bold"), fg="#0F172A", bg="#FFFFFF")
            self.lbl_title.pack(anchor="w")
            self.lbl_sub = tk.Label(self.card, text="High-Speed Arabic & English Document & OCR Converter", font=("Segoe UI", 11), fg="#64748B", bg="#FFFFFF")
            self.lbl_sub.pack(anchor="w", pady=(2, 16))

            in_box = tk.Frame(self.card, bg="#FFFFFF")
            in_box.pack(fill="x", pady=6)
            tk.Label(in_box, text="Input File:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF").pack(anchor="w")
            self.txt_in = tk.Entry(in_box, font=("Segoe UI", 10))
            self.txt_in.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
            tk.Button(in_box, text="Browse", command=self.browse_in, bg="#2563EB", fg="#FFFFFF", relief="flat", padx=10).pack(side="right")

            fmt_box = tk.Frame(self.card, bg="#FFFFFF")
            fmt_box.pack(fill="x", pady=6)
            tk.Label(fmt_box, text="Format:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF").pack(anchor="w")
            self.cmb_fmt = ttk.Combobox(fmt_box, values=FORMAT_OPTIONS, state="readonly", font=("Segoe UI", 10))
            self.cmb_fmt.set(FORMAT_OPTIONS[0])
            self.cmb_fmt.pack(fill="x")

            out_box = tk.Frame(self.card, bg="#FFFFFF")
            out_box.pack(fill="x", pady=6)
            tk.Label(out_box, text="Output File:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF").pack(anchor="w")
            self.txt_out = tk.Entry(out_box, font=("Segoe UI", 10))
            self.txt_out.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
            tk.Button(out_box, text="Change", command=self.browse_out, bg="#E2E8F0", relief="flat", padx=10).pack(side="right")

            self.btn_run = tk.Button(self.card, text="Run OCR & Export", font=("Segoe UI", 11, "bold"), bg="#2563EB", fg="#FFFFFF", relief="flat", pady=8, command=self.start_conv)
            self.btn_run.pack(fill="x", pady=14)

            self.prog = ttk.Progressbar(self.card, orient="horizontal", mode="determinate")
            self.prog.pack(fill="x", pady=4)
            self.lbl_st = tk.Label(self.card, text="Ready", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#475569")
            self.lbl_st.pack(anchor="w")
            self.lbl_dt = tk.Label(self.card, text="", font=("Segoe UI", 9), bg="#FFFFFF", fg="#64748B")
            self.lbl_dt.pack(anchor="w")

        def browse_in(self):
            f = filedialog.askopenfilename(filetypes=FILE_TYPES)
            if f:
                self.txt_in.delete(0, "end")
                self.txt_in.insert(0, f)
                ext = FORMAT_CONFIG[self.cmb_fmt.get()]["ext"]
                self.txt_out.delete(0, "end")
                self.txt_out.insert(0, os.path.splitext(f)[0] + ext)

        def browse_out(self):
            f = filedialog.asksaveasfilename()
            if f:
                self.txt_out.delete(0, "end")
                self.txt_out.insert(0, f)

        def update_status_safe(self, message=None, progress=None, detail=None, text_color=None):
            def _apply():
                if message: self.lbl_st.configure(text=message)
                if detail: self.lbl_dt.configure(text=detail)
                if progress is not None: self.prog["value"] = progress * 100
                if text_color: self.lbl_st.configure(fg=text_color)
            self.after(0, _apply)

        def set_ui_state_processing(self, val):
            self.btn_run.configure(state="disabled" if val else "normal")

        def show_action_buttons(self): pass
        def hide_action_buttons(self): pass

        def start_conv(self):
            inp = self.txt_in.get().strip()
            outp = self.txt_out.get().strip()
            fmt = self.cmb_fmt.get().strip()
            if not inp or not os.path.exists(inp):
                messagebox.showerror("Error", "Please select a valid input document.")
                return
            self.set_ui_state_processing(True)
            threading.Thread(target=self._run_conversion_worker, args=(inp, outp, fmt), daemon=True).start()

        def _open_page_1_preview(self, page_1_info, total_pages, target_format, on_decision):
            try:
                PagePreviewDialog(self, page_1_info, total_pages, target_format, on_decision)
            except Exception as e:
                on_decision(True)

        def _on_conversion_cancelled(self):
            self.set_ui_state_processing(False)




def main():
    multiprocessing.freeze_support()
    app = PDF2textApp()
    app.mainloop()


if __name__ == "__main__":
    main()
