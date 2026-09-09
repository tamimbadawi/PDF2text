"""
PDF2text - Desktop Document & OCR Converter
Converts PDF, scanned documents, and images to Excel (.xlsx), Word (.docx), and Text (.txt).
Features structured table grid extraction, OpenCV morphology line detection, and bounding-box spatial clustering.
"""

import os
import re
import sys
import json
import difflib
import threading
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
import time
import warnings
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# PyTorch import and check
try:
    import torch
    HAS_TORCH = True
except ImportError:
    HAS_TORCH = False

# Suppress PyTorch and EasyOCR deprecation warnings
warnings.filterwarnings("ignore", category=UserWarning)

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

# Document, Computer Vision & OCR processing libraries
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
    import easyocr
    HAS_EASYOCR = True
except ImportError:
    HAS_EASYOCR = False

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

try:
    from markitdown import MarkItDown
    HAS_MARKITDOWN = True
except ImportError:
    HAS_MARKITDOWN = False

try:
    import surya
    import surya.layout
    import surya.table_rec
    HAS_SURYA = True
except Exception:
    HAS_SURYA = False


# Supported formats & configuration
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
    ("All Supported Documents", "*.pdf;*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tiff;*.docx;*.pptx;*.xlsx;*.xls;*.txt;*.html"),
    ("PDF Documents (*.pdf)", "*.pdf"),
    ("Images (*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tiff)", "*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tiff"),
    ("Word Documents (*.docx)", "*.docx"),
    ("PowerPoint Presentations (*.pptx)", "*.pptx"),
    ("Excel Spreadsheets (*.xlsx;*.xls)", "*.xlsx;*.xls"),
    ("Text & Web Files (*.txt;*.html;*.csv;*.json;*.xml)", "*.txt;*.html;*.csv;*.json;*.xml"),
    ("All Files (*.*)", "*.*")
]


# ==============================================================================
# Egyptian Arabic Construction Dictionary & Spell-Correction Engine
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
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        paths_to_try = [
            dict_path,
            os.path.join(base_dir, "dictionary.json"),
            os.path.join(os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)), "dictionary.json"),
            os.path.abspath("dictionary.json")
        ]
        loaded = False
        for p in paths_to_try:
            if p and os.path.isfile(p):
                try:
                    with open(p, "r", encoding="utf-8") as f:
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
                    break
                except Exception as e:
                    print(f"Warning: Failed to parse dictionary at {p}: {e}")

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
                matches = difflib.get_close_matches(core, self.known_words, n=1, cutoff=0.86)
                if matches:
                    core = matches[0]

            corrected_tokens.append(prefix + core + suffix)

        result = " ".join(corrected_tokens)

        # 3. Clean up dimensions & standard expressions
        result = re.sub(r"(\d+)\s*['’`*xX×]\s*(\d+)\s*س?م?", r"\1*\2 سم", result)
        result = re.sub(r"(\d+)\s+(\d+)\s*س\b", r"\1*\2 سم", result)
        result = re.sub(r"(\d+)\s+(\d+)\s*سم\b", r"\1*\2 سم", result)
        result = re.sub(r"سمعمق", "سم عمق", result)
        result = re.sub(r"\s+", " ", result).strip()

        return result

    def correct_unit(self, unit: str) -> str:
        if not unit:
            return ""
        u = unit.strip().replace(" ", "").replace(".", "")
        if any(k in u for k in ["عدد", "عدذ", "عذد", "عد"]):
            return "عدد"
        if any(k in u for k in ["مط", "م.ط", "ط", "درطط", "م"]):
            return "م.ط"
        return unit.strip()


# ==============================================================================
# Structured Table Extraction Engine
# ==============================================================================
class StructuredTableExtractor:
    """
    Extracts structured multi-column table grids from document page images
    and OCR bounding boxes using spatial bounding-box clustering,
    RTL text and dimension sequence alignment, explicit 3-column mapping,
    and section category header merging.
    """

    @staticmethod
    def is_arabic(text: str) -> bool:
        return any('\u0600' <= ch <= '\u06FF' for ch in str(text))

    @staticmethod
    def clean_arabic_dimensions(s: str) -> str:
        """Fixes scrambled dimensions, inverted English/Arabic nouns, and preserves LTR numeric sequences."""
        if not s:
            return ""

        # 1. Invert English technical acronyms prefixing Arabic nouns (e.g. "HDPE مواسير" -> "مواسير HDPE")
        s = re.sub(r"^([A-Za-z0-9\-_]{2,})\s+([\u0600-\u06FF]+)", r"\2 \1", s)
        s = re.sub(r"\b([A-Za-z0-9\-_]{2,})\s+(مواسير|محبس|طوب|صرف|غرفة|سخان|كرفان|الواح|زجاج)\b", r"\2 \1", s)

        # 2. 3-part dimensions (e.g. 20"20"40 or 40*20*10 or 20 20 40)
        s = re.sub(r"(\d+)\s*[\'’`\"*xX×*٢،,]\s*(\d+)\s*[\'’`\"*xX×*٢،,]\s*(\d+)\s*(?:س?م|م?م|م)?", r"\1*\2*\3 سم", s)

        # 3. 2-part dimensions (e.g. 60*90 سم, 60'90سم, 90٢60سم, 60 90 سم, 25'25 سم)
        s = re.sub(r"(\d+)\s*[\'’`\"*xX×*٢،,]\s*(\d+)\s*(س?م|م?م|م)?\b", r"\1*\2 سم", s)
        s = re.sub(r"(\d+)\s+(\d+)\s*س\b", r"\1*\2 سم", s)
        s = re.sub(r"(\d+)\s+(\d+)\s*سم\b", r"\1*\2 سم", s)

        # 4. Dimension with depth keyword (e.g. سمعمق 12 -> سم عمق 120 سم or عمق 120)
        s = re.sub(r"سمعمق", "سم عمق", s)
        s = re.sub(r"عمق\s*(\d+)\s*(?:س?م)?\b", r"عمق \1 سم", s)

        # 5. Fix common unit notations (e.g. 25مع -> 25 مم, 75م~ -> 75 مم, 160مع -> 160 مم)
        s = re.sub(r"(\d+)\s*مع\b", r"\1 مم", s)
        s = re.sub(r"(\d+)\s*م~\b", r"\1 مم", s)
        s = re.sub(r"(\d+)\s*م~~\b", r"\1 مم", s)

        # 6. Normalize duplicate spaces and strip
        s = re.sub(r"\s+", " ", s).strip()
        return s

    @classmethod
    def assemble_cell_text(cls, items_list, is_rtl=True):
        """Assembles multiple OCR detections within a single table cell in natural reading order."""
        if not items_list:
            return ""
        if len(items_list) == 1:
            t = items_list[0]['text']
            return cls.clean_arabic_dimensions(t) if is_rtl else t

        # Group text boxes into horizontal sub-lines by Y proximity
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
            # In RTL (Arabic), sort horizontally by descending xc (rightmost piece first)
            # In LTR (English), sort horizontally by ascending xc (leftmost piece first)
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
        """Normalizes unit abbreviations in Arabic table cells."""
        if not u:
            return ""
        u_clean = u.strip().replace(" ", "").replace(".", "")
        if any(k in u_clean for k in ["عدد", "عدذ", "عذد", "عد"]):
            return "عدد"
        if any(k in u_clean for k in ["مط", "م.ط", "ط", "درطط", "م"]):
            return "م.ط"
        return u.strip()

    @classmethod
    def extract_table_from_page(cls, img_bgr, ocr_results, reader=None, apply_dictionary=True):
        """
        Extracts a clean, structured table from page image and EasyOCR bounding boxes using
        OpenCV morphology grid detection, strict vertical row clustering, explicit RTL column mapping,
        and targeted cell crop recovery.
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

        # Determine dimensions
        if img_bgr is not None:
            img_h, img_w = img_bgr.shape[:2]
        else:
            all_xs = [pt[0] for bbox, _, _ in ocr_results for pt in bbox]
            all_ys = [pt[1] for bbox, _, _ in ocr_results for pt in bbox]
            img_w = max(all_xs) + 50 if all_xs else 1653
            img_h = max(all_ys) + 50 if all_ys else 2339

        # 1. Standardize OCR items & filter vertical grid-line noise artifacts
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

            # Filter vertical grid lines & CRAFT merged line artifacts
            if h > 45 and float(conf) < 0.25:
                continue
            if h > 70 and w < 60:
                continue
            if h > (img_h * 0.025) and w < (img_w * 0.04) and (float(conf) < 0.35 or t in (":", ".", "|", "", "-", "=", "ة")):
                continue
            if float(conf) < 0.15 and not any(c.isalnum() for c in t):
                continue

            items.append({
                'x0': min(xs), 'x1': max(xs),
                'y0': min(ys), 'y1': max(ys),
                'xc': sum(xs) / len(xs),
                'yc': sum(ys) / len(ys),
                'h': h, 'w': w,
                'text': t, 'conf': float(conf)
            })

        is_rtl = (arabic_chars / max(1, total_chars)) > 0.25

        # 2. Horizontal Grid Line Detection via OpenCV morphology
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

        # 3. Form vertical row bands
        x_price_thresh = img_w * 0.24
        x_unit_thresh = img_w * 0.42

        title_headers = []
        footer_notes = []
        table_bands = []  # list of (row_items, y_top, y_bot)

        if has_physical_grid:
            first_grid_y = line_ys[0]
            last_grid_y = line_ys[-1]

            # Title banner above first grid line
            top_items = [it for it in items if it['yc'] < first_grid_y]
            if top_items:
                title_headers.append(cls.clean_arabic_dimensions(" ".join(x['text'] for x in sorted(top_items, key=lambda x: x['xc'], reverse=True))))

            # Footer notes below last grid line
            bot_items = [it for it in items if it['yc'] > last_grid_y]
            if bot_items:
                footer_notes.append(" ".join(x['text'] for x in bot_items))

            # Split unusually tall grid bands (e.g. table header + first banner)
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
            # Fallback: strict vertical spatial clustering
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
                if y_t > (img_h * 0.95):
                    footer_notes.append(cls.assemble_cell_text(r, is_rtl))
                    continue
                table_bands.append((r, y_t, y_b))

        # 4. Map columns & detect section category headers
        raw_table_rows = []
        raw_category_rows = set()

        for r_items, y_top, y_bot in table_bands:
            price_items = [x for x in r_items if x['xc'] < x_price_thresh]
            unit_items = [x for x in r_items if x_price_thresh <= x['xc'] < x_unit_thresh]
            desc_items = [x for x in r_items if x['xc'] >= x_unit_thresh]

            row_text = " ".join(x['text'] for x in sorted(r_items, key=lambda x: x['xc'], reverse=True))
            row_x0 = min((x['x0'] for x in r_items), default=0)
            row_x1 = max((x['x1'] for x in r_items), default=0)
            row_xc = (row_x0 + row_x1) / 2 if r_items else 0

            is_cat = False
            is_header = False

            if any("البند" in x['text'] for x in desc_items) or any("القياس" in x['text'] for x in unit_items):
                is_header = True
            elif any(k in row_text for k in ["مصنعية الغرف", "السخانات", "تغذية", "حائطي بول", "حانطى بول", "UPVC", "سيملس"]):
                if not any(k in row_text for k in ["قطر خارج", "قطر 25", "قطر 32", "قطر 50", "بوصة", "بوصا"]):
                    is_cat = True
            elif not price_items and not unit_items:
                if (img_w * 0.28 <= row_xc <= img_w * 0.72) and (y_bot - y_top >= 26):
                    if not any(k in row_text for k in ["قطر", "مقاس", "مفاس", "بوصة", "بوصا"]):
                        is_cat = True

            r_idx = len(raw_table_rows)
            if is_cat:
                raw_category_rows.add(r_idx)
                cat_desc = cls.clean_arabic_dimensions(row_text)
                raw_table_rows.append({
                    'is_category': True,
                    'is_header': False,
                    'desc': cat_desc,
                    'unit': '',
                    'price': '',
                    'y0': y_top, 'y1': y_bot
                })
            elif is_header:
                raw_table_rows.append({
                    'is_category': False,
                    'is_header': True,
                    'desc': "البند",
                    'unit': "وحدة القياس",
                    'price': "متوسط السعر",
                    'y0': y_top, 'y1': y_bot
                })
            else:
                desc_str = cls.clean_arabic_dimensions(" ".join(x['text'] for x in sorted(desc_items, key=lambda x: x['xc'], reverse=True)))
                unit_str = cls.clean_unit(" ".join(x['text'] for x in unit_items))
                price_str = " ".join(x['text'] for x in sorted(price_items, key=lambda x: x['xc'])).strip()
                raw_table_rows.append({
                    'is_category': False,
                    'is_header': False,
                    'desc': desc_str,
                    'unit': unit_str,
                    'price': price_str,
                    'y0': y_top, 'y1': y_bot
                })

        # 5. Targeted crop recovery for missing prices/units/descriptions
        if reader is not None and img_bgr is not None:
            for idx, r_data in enumerate(raw_table_rows):
                if r_data['is_category'] or r_data['is_header']:
                    continue
                y0 = max(0, int(r_data['y0']) - 2)
                y1 = min(img_h, int(r_data['y1']) + 2)

                # Recover missing price
                if not r_data['price'] or not any(ch.isdigit() for ch in r_data['price']):
                    price_crop = img_bgr[y0:y1, int(img_w * 0.04):int(x_price_thresh)]
                    if price_crop.shape[0] > 5 and price_crop.shape[1] > 10:
                        try:
                            res_p = reader.readtext(price_crop, detail=1)
                            p_nums = [it[1] for it in res_p if any(c.isdigit() for c in it[1])]
                            if p_nums:
                                r_data['price'] = " ".join(p_nums).strip()
                        except Exception:
                            pass

                # Recover missing unit
                if not r_data['unit']:
                    unit_crop = img_bgr[y0:y1, int(x_price_thresh):int(x_unit_thresh)]
                    if unit_crop.shape[0] > 5 and unit_crop.shape[1] > 10:
                        try:
                            res_u = reader.readtext(unit_crop, detail=1)
                            u_texts = [cls.clean_unit(it[1]) for it in res_u if cls.clean_unit(it[1])]
                            if u_texts:
                                r_data['unit'] = u_texts[0]
                        except Exception:
                            pass

                # Recover incomplete description (e.g. solitary 'مقاس' missing dimensions)
                desc = r_data['desc']
                if len(desc) < 7 or not any(ch.isdigit() for ch in desc):
                    desc_crop = img_bgr[y0:y1, int(x_unit_thresh):int(img_w * 0.98)]
                    if desc_crop.shape[0] > 5 and desc_crop.shape[1] > 10:
                        try:
                            res_d = reader.readtext(desc_crop, detail=1)
                            d_texts = [it[1] for it in sorted(res_d, key=lambda x: (x[0][0][0] + x[0][1][0]) / 2, reverse=True)]
                            if d_texts:
                                new_desc = cls.clean_arabic_dimensions(" ".join(d_texts)).strip()
                                if any(ch.isdigit() for ch in new_desc) or len(new_desc) > len(desc):
                                    r_data['desc'] = new_desc
                        except Exception:
                            pass

        # 6. Contextual Unit Inheritance across section items
        active_unit = "عدد"
        for idx, r_data in enumerate(raw_table_rows):
            if r_data['is_category']:
                desc_lower = r_data['desc']
                if any(k in desc_lower for k in ["مواسير", "صرف", "UPVC", "سيملس", "HDPE", "بولي"]):
                    active_unit = "م.ط"
                else:
                    active_unit = "عدد"
            elif not r_data['is_header']:
                if not r_data['unit']:
                    r_data['unit'] = active_unit
                else:
                    active_unit = r_data['unit']

        # 7. Apply Egyptian Arabic Dictionary & Spell Correction if enabled
        if apply_dictionary:
            corrector = ArabicDictionaryCorrector.get_instance()
            title_headers = [corrector.correct_text(h) for h in title_headers]
            for r in raw_table_rows:
                if not r['is_header']:
                    r['desc'] = corrector.correct_text(r['desc'])
                    if not r['is_category']:
                        r['unit'] = corrector.correct_unit(r['unit'])

        final_rows = [[r['desc'], r['unit'], r['price']] for r in raw_table_rows]

        return {
            'title_headers': title_headers,
            'table_rows': final_rows,
            'category_rows': raw_category_rows,
            'footer_notes': footer_notes,
            'is_rtl': is_rtl,
            'has_grid': has_physical_grid
        }


# ==============================================================================
# Model & Dictionary Path Resolvers
# ==============================================================================
def get_model_storage_dir():
    """Locates offline bundled model weights or falls back to user cache."""
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    m_dir = os.path.join(base_dir, "models")
    if os.path.isdir(m_dir) and os.path.isfile(os.path.join(m_dir, "craft_mlt_25k.pth")):
        return m_dir

    exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    m_dir2 = os.path.join(exe_dir, "models")
    if os.path.isdir(m_dir2) and os.path.isfile(os.path.join(m_dir2, "craft_mlt_25k.pth")):
        return m_dir2

    m_dir3 = os.path.join(exe_dir, "_internal", "models")
    if os.path.isdir(m_dir3) and os.path.isfile(os.path.join(m_dir3, "craft_mlt_25k.pth")):
        return m_dir3

    return os.path.join(os.path.expanduser("~"), ".EasyOCR", "model")


def get_dictionary_path():
    """Locates dictionary.json in bundled directory or alongside executable."""
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    p1 = os.path.join(base_dir, "dictionary.json")
    if os.path.isfile(p1):
        return p1
    exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    p2 = os.path.join(exe_dir, "dictionary.json")
    if os.path.isfile(p2):
        return p2
    p3 = os.path.join(exe_dir, "_internal", "dictionary.json")
    if os.path.isfile(p3):
        return p3
    return os.path.abspath("dictionary.json")


# ==============================================================================
# High-Performance Multiprocessing Worker Entry Points
# ==============================================================================
_GLOBAL_WORKER_READER = None
_GLOBAL_WORKER_CORRECTOR = None


def init_worker_process(models_dir: str = None, dict_path: str = None):
    """
    Initializes EasyOCR reader and Egyptian Arabic dictionary once per child process.
    Pre-warms CRAFT detector and CRNN recognizer with thread limits to prevent CPU oversubscription.
    """
    global _GLOBAL_WORKER_READER, _GLOBAL_WORKER_CORRECTOR
    try:
        if HAS_TORCH:
            import torch
            torch.set_num_threads(2)
        if models_dir is None:
            models_dir = get_model_storage_dir()
        if HAS_EASYOCR and _GLOBAL_WORKER_READER is None:
            _GLOBAL_WORKER_READER = easyocr.Reader(
                ['ar', 'en'],
                gpu=False,
                verbose=False,
                model_storage_directory=models_dir,
                download_enabled=False
            )
        if dict_path is None:
            dict_path = get_dictionary_path()
        _GLOBAL_WORKER_CORRECTOR = ArabicDictionaryCorrector.get_instance(dict_path)
    except Exception as e:
        print(f"[Worker Init] Failed to initialize worker process: {e}", file=sys.stderr)


def process_pdf_page_job(page_idx: int, pdf_path: str, models_dir: str = None, dict_path: str = None, apply_dictionary: bool = True, dpi: int = 200):
    """
    Picklable multiprocessing job for parsing a single PDF page:
    1. Direct zero-copy rasterization via PyMuPDF samples buffer.
    2. Batched EasyOCR inference (batch_size=8) inside torch.inference_mode().
    3. Structured spatial table extraction, column mapping, and merged row detection.
    4. Egyptian dictionary spell-correction.
    """
    global _GLOBAL_WORKER_READER, _GLOBAL_WORKER_CORRECTOR
    if _GLOBAL_WORKER_READER is None:
        init_worker_process(models_dir, dict_path)

    import pymupdf
    import numpy as np

    doc = pymupdf.open(pdf_path)
    try:
        page = doc[page_idx]
        pix = page.get_pixmap(dpi=dpi)
        raw = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

        # Zero-copy color space conversions
        if pix.n == 4:
            img_bgr = cv2.cvtColor(raw, cv2.COLOR_RGBA2BGR) if HAS_CV2 else None
            img_ocr = cv2.cvtColor(raw, cv2.COLOR_RGBA2RGB) if HAS_CV2 else raw
        elif pix.n == 3:
            img_bgr = cv2.cvtColor(raw, cv2.COLOR_RGB2BGR) if HAS_CV2 else None
            img_ocr = raw
        else:
            img_bgr = raw
            img_ocr = raw

        # Batched inference with torch.inference_mode()
        if HAS_TORCH:
            import torch
            with torch.inference_mode():
                ocr_results = _GLOBAL_WORKER_READER.readtext(
                    img_ocr,
                    batch_size=8,
                    detail=1,
                    paragraph=False
                )
        else:
            ocr_results = _GLOBAL_WORKER_READER.readtext(
                img_ocr,
                batch_size=8,
                detail=1,
                paragraph=False
            )

        # Spatial table extraction
        table_data = StructuredTableExtractor.extract_table_from_page(
            img_bgr,
            ocr_results,
            reader=_GLOBAL_WORKER_READER,
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
            'table_data': {'title_headers': [], 'table_rows': [], 'category_rows': set(), 'has_grid': False},
            'ocr_results': [],
            'block_count': 0,
            'char_count': 0,
            'error': str(e)
        }
    finally:
        doc.close()


# ==============================================================================
# Base Conversion Mixin
# ==============================================================================
class ConversionEngineMixin:
    """Shared document parsing, OCR pipeline, and multi-column file export helpers."""

    def __init__(self):
        self.ocr_reader = None
        self.markitdown_converter = MarkItDown() if HAS_MARKITDOWN else None
        self.last_converted_path = None

    def _get_model_storage_dir(self):
        """Locates offline bundled model weights or falls back to user cache."""
        return get_model_storage_dir()

    def _get_ocr_reader(self):
        """Initializes or returns cached EasyOCR reader for Arabic and English."""
        if self.ocr_reader is None:
            if not HAS_EASYOCR:
                raise RuntimeError("EasyOCR is not installed. Please install it via 'pip install easyocr'.")
            model_dir = self._get_model_storage_dir()
            self.ocr_reader = easyocr.Reader(
                ["ar", "en"],
                gpu=False,
                verbose=False,
                model_storage_directory=model_dir
            )
        return self.ocr_reader

    def _save_excel(self, extracted_pages_data, output_path, input_path):
        """Generates a structured multi-column Excel workbook with aligned table columns and headers."""
        if not HAS_OPENPYXL:
            raise RuntimeError("openpyxl is not installed. Please install it via 'pip install openpyxl'.")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default blank sheet

        title_font = Font(name="Segoe UI", size=13, bold=True, color="1E3A8A")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")

        # Yellow banner fill matching the original document
        cat_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        cat_font = Font(name="Segoe UI", size=11, bold=True, color="1E293B")

        alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        data_font = Font(name="Segoe UI", size=10, color="0F172A")

        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        right_align = Alignment(horizontal="right", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        thin_side = Side(style="thin", color="CBD5E1")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # Seamless merged banner borders without interior vertical lines
        cat_border_right = Border(top=thin_side, bottom=thin_side, right=thin_side)
        cat_border_mid = Border(top=thin_side, bottom=thin_side)
        cat_border_left = Border(top=thin_side, bottom=thin_side, left=thin_side)

        # 1. Consolidated "All Extracted Data" sheet
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
            is_rtl = t_data['is_rtl']
            align_text = right_align if is_rtl else left_align
            category_rows = t_data.get('category_rows', set())

            # Page header banner
            ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
            banner = ws_all.cell(row=current_row, column=1)
            banner.value = f"--- [ صفحة / Page {page_num} ] ---"
            banner.font = Font(name="Segoe UI", size=11, bold=True, color="475569")
            banner.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            banner.alignment = center_align
            ws_all.row_dimensions[current_row].height = 28
            current_row += 1

            # Title headers above the table
            for h_text in t_data['title_headers']:
                ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                c = ws_all.cell(row=current_row, column=1)
                c.value = h_text
                c.font = title_font
                c.alignment = center_align
                ws_all.row_dimensions[current_row].height = 26
                current_row += 1

            # Table rows mapped across explicit columns: Col 1=Desc, Col 2=Unit, Col 3=Price
            for r_i, row in enumerate(t_data['table_rows']):
                is_first_row = (r_i == 0)
                is_category = (r_i in category_rows)

                if is_category:
                    # Full-width section category banner merged across A to C
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
                    # Data columns: Col 1: Description, Col 2: Unit, Col 3: Price
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
                            if c_i in (1, 2):  # Unit and Price centered
                                c.alignment = center_align
                            else:
                                c.alignment = align_text

                    ws_all.row_dimensions[current_row].height = 26 if is_first_row else 22

                current_row += 1

            # Footers
            for f_text in t_data['footer_notes']:
                ws_all.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                c = ws_all.cell(row=current_row, column=1)
                c.value = f_text
                c.font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
                c.alignment = center_align
                current_row += 1

            current_row += 1  # Spacing between pages

        # Column widths for consolidated sheet
        ws_all.column_dimensions["A"].width = 52
        ws_all.column_dimensions["B"].width = 16
        ws_all.column_dimensions["C"].width = 18

        # 2. Individual Per-Page Sheets for exact table layout
        for page_info in extracted_pages_data:
            page_num = page_info['page_num']
            t_data = page_info['table_data']
            is_rtl = t_data['is_rtl']
            align_text = right_align if is_rtl else left_align
            category_rows = t_data.get('category_rows', set())

            ws_page = wb.create_sheet(title=f"Page {page_num}")
            if is_rtl:
                try:
                    ws_page.sheet_view.rightToLeft = True
                except Exception:
                    pass

            p_row = 1

            # Title headers
            for h_text in t_data['title_headers']:
                ws_page.merge_cells(start_row=p_row, start_column=1, end_row=p_row, end_column=3)
                c = ws_page.cell(row=p_row, column=1)
                c.value = h_text
                c.font = title_font
                c.alignment = center_align
                ws_page.row_dimensions[p_row].height = 26
                p_row += 1

            # Table rows
            for r_i, row in enumerate(t_data['table_rows']):
                is_first_row = (r_i == 0)
                is_category = (r_i in category_rows)

                if is_category:
                    # Full-width section category banner merged across A to C
                    ws_page.merge_cells(start_row=p_row, start_column=1, end_row=p_row, end_column=3)
                    c1 = ws_page.cell(row=p_row, column=1)
                    c1.value = row[0]
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
                        c = ws_page.cell(row=p_row, column=col_idx)
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
                            if c_i in (1, 2):  # Unit and Price centered
                                c.alignment = center_align
                            else:
                                c.alignment = align_text

                    ws_page.row_dimensions[p_row].height = 26 if is_first_row else 22

                p_row += 1

            # Footers
            for f_text in t_data['footer_notes']:
                ws_page.merge_cells(start_row=p_row, start_column=1, end_row=p_row, end_column=3)
                c = ws_page.cell(row=p_row, column=1, value=f_text)
                c.font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
                c.alignment = center_align
                p_row += 1

            # Column widths for per-page sheet
            ws_page.column_dimensions["A"].width = 54
            ws_page.column_dimensions["B"].width = 16
            ws_page.column_dimensions["C"].width = 18

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        wb.save(output_path)

    def _save_word(self, extracted_pages_data, output_path, input_path):
        """Generates a structured Word document (.docx) with native Word tables and merged section headers."""
        if not HAS_DOCX:
            raise RuntimeError("python-docx is not installed. Please install it via 'pip install python-docx'.")

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
            category_rows = t_data.get('category_rows', set())

            doc.add_heading(f"Page {page_num} / صفحة {page_num}", level=2)

            for h_text in t_data['title_headers']:
                p = doc.add_paragraph(h_text)
                p.runs[0].font.bold = True

            rows = t_data['table_rows']
            if rows:
                num_cols = max(len(r) for r in rows)
                table = doc.add_table(rows=len(rows), cols=num_cols)
                table.style = 'Table Grid'

                for r_i, row in enumerate(rows):
                    is_cat = (r_i in category_rows) or (len([c for c in row if c]) == 1 and len(row) > 1 and r_i > 0)
                    if is_cat and num_cols > 1:
                        cell_start = table.cell(r_i, 0)
                        cell_end = table.cell(r_i, num_cols - 1)
                        merged_cell = cell_start.merge(cell_end)
                        merged_cell.text = row[0]
                        for p in merged_cell.paragraphs:
                            for run in p.runs:
                                run.font.bold = True
                    else:
                        for c_i, cell_text in enumerate(row):
                            cell = table.cell(r_i, c_i)
                            cell.text = cell_text
                            if r_i == 0:
                                for p in cell.paragraphs:
                                    for run in p.runs:
                                        run.font.bold = True
            else:
                p = doc.add_paragraph("[No table structure detected on this page]")
                p.runs[0].font.italic = True

            for f_text in t_data['footer_notes']:
                p = doc.add_paragraph(f_text)
                p.runs[0].font.italic = True

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        doc.save(output_path)

    def _save_text(self, extracted_pages_data, output_path, input_path):
        """Generates clean UTF-8 text with tab-delimited columns for spreadsheet copy-pasting."""
        total_pages = len(extracted_pages_data)

        lines = [
            "=" * 70,
            "PDF2text - Extracted Structured Document",
            f"Source File   : {os.path.basename(input_path)}",
            f"Total Pages   : {total_pages}",
            f"Generated At  : {time.strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 70,
            ""
        ]

        for page_info in extracted_pages_data:
            page_num = page_info['page_num']
            t_data = page_info['table_data']

            lines.append(f"--- [ صفحة / Page {page_num} ] ---")
            lines.append("")

            for h_text in t_data['title_headers']:
                lines.append(f"# {h_text}")

            for row in t_data['table_rows']:
                lines.append("\t".join(row))

            for f_text in t_data['footer_notes']:
                lines.append(f"// {f_text}")

            lines.append("")

        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")


# ==============================================================================
# Interactive First-Page Preview Dialog
# ==============================================================================
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


# ==============================================================================
# Modern CustomTkinter Application Implementation
# ==============================================================================
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
            self.update_status_safe(
                message="Operation cancelled by user.",
                progress=0.0,
                detail="No files were exported. Ready for next document.",
                text_color="#64748B"
            )

        def _run_conversion_worker(self, input_path: str, output_path: str, target_format: str):
            start_time = time.time()
            try:
                # Step 1: Initialize or verify OCR engine
                self.update_status_safe(
                    message="[1/4] Preparing OCR engine (Arabic & English)...",
                    progress=0.05,
                    detail="Loading cached EasyOCR weights...",
                    text_color="#2563EB"
                )
                reader = self._get_ocr_reader()

                # Step 2: Open and inspect input document
                ext = os.path.splitext(input_path)[1].lower()
                self.update_status_safe(
                    message=f"[2/4] Opening document: {os.path.basename(input_path)}...",
                    progress=0.12,
                    detail="Analyzing document structure...",
                    text_color="#2563EB"
                )

                extracted_pages_data = []  # List of page data dicts
                total_blocks = 0
                total_chars = 0

                if ext == ".pdf":
                    if not HAS_PYMUPDF:
                        raise RuntimeError("PyMuPDF is required to process PDF documents. Run 'pip install pymupdf'.")

                    doc = pymupdf.open(input_path)
                    total_pages = len(doc)
                    self.update_status_safe(
                        message=f"[2/4] Document loaded: {total_pages} page(s) detected.",
                        progress=0.15,
                        detail="Rendering Page 1 for interactive preview & verification...",
                        text_color="#2563EB"
                    )

                    # Step 3A: Process Page 1 with Zero-Copy Rasterization & Batched Inference
                    self.update_status_safe(
                        message=f"[3/4] Rendering Page 1 of {total_pages}...",
                        progress=0.18,
                        detail="Zero-copy rasterization at 200 DPI...",
                        text_color="#2563EB"
                    )
                    page0 = doc[0]
                    pix0 = page0.get_pixmap(dpi=200)
                    raw0 = np.frombuffer(pix0.samples, dtype=np.uint8).reshape(pix0.h, pix0.w, pix0.n)
                    if pix0.n == 4:
                        img_bgr0 = cv2.cvtColor(raw0, cv2.COLOR_RGBA2BGR) if HAS_CV2 else None
                        img_ocr0 = cv2.cvtColor(raw0, cv2.COLOR_RGBA2RGB) if HAS_CV2 else raw0
                    elif pix0.n == 3:
                        img_bgr0 = cv2.cvtColor(raw0, cv2.COLOR_RGB2BGR) if HAS_CV2 else None
                        img_ocr0 = raw0
                    else:
                        img_bgr0 = raw0
                        img_ocr0 = raw0

                    self.update_status_safe(
                        message=f"[3/4] Running EasyOCR on Page 1 of {total_pages}...",
                        progress=0.22,
                        detail="Batched neural inference in torch.inference_mode()...",
                        text_color="#2563EB"
                    )
                    if HAS_TORCH:
                        with torch.inference_mode():
                            ocr_res0 = reader.readtext(img_ocr0, batch_size=8, detail=1, paragraph=False)
                    else:
                        ocr_res0 = reader.readtext(img_ocr0, batch_size=8, detail=1, paragraph=False)

                    self.update_status_safe(
                        message=f"[3/4] Extracting structured table grid on Page 1...",
                        progress=0.25,
                        detail="Applying spatial clustering, RTL alignment & category merging...",
                        text_color="#2563EB"
                    )
                    table_data0 = StructuredTableExtractor.extract_table_from_page(img_bgr0, ocr_res0, reader=reader, apply_dictionary=True)
                    page_1_info = {
                        'page_num': 1,
                        'table_data': table_data0,
                        'ocr_results': ocr_res0
                    }
                    extracted_pages_data.append(page_1_info)
                    total_blocks += len(ocr_res0)
                    total_chars += sum(len(str(it[1])) for it in ocr_res0)

                    # Launch Background Prefetch Pipeline for Pages 2..N
                    pool_executor = None
                    prefetch_futures = {}
                    prefetch_results = {}
                    prefetch_cancelled = threading.Event()
                    num_workers = min(4, max(1, os.cpu_count() // 2))

                    if total_pages > 1:
                        models_dir = self._get_model_storage_dir()
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
                            print(f"Notice: ProcessPoolExecutor fallback: {pool_err}")
                            pool_executor = None

                    # Trigger interactive preview modal
                    preview_event = threading.Event()
                    preview_result = {'accepted': False}

                    def _on_preview_choice(choice):
                        if isinstance(choice, dict):
                            preview_result.update(choice)
                        else:
                            preview_result['accepted'] = bool(choice)
                        preview_event.set()

                    self.update_status_safe(
                        message=f"[Preview] Page 1 ready ({len(table_data0['table_rows'])} rows). Review while remaining pages prefetch...",
                        progress=0.28,
                        detail=f"Parallel pipeline active across {num_workers} CPU cores. Review Page 1 table layout.",
                        text_color="#D97706"
                    )
                    self.after(0, self._open_page_1_preview, page_1_info, total_pages, target_format, _on_preview_choice)

                    # Block worker thread until user decides in the preview dialog
                    preview_event.wait()

                    if not preview_result.get('accepted', False):
                        prefetch_cancelled.set()
                        if pool_executor:
                            pool_executor.shutdown(wait=False, cancel_futures=True)
                        doc.close()
                        self.after(0, self._on_conversion_cancelled)
                        return

                    # Update page 1 with any user edits made in preview
                    if 'table_rows' in preview_result and preview_result['table_rows']:
                        page_1_info['table_data']['table_rows'] = preview_result['table_rows']
                    if 'category_rows' in preview_result and preview_result['category_rows'] is not None:
                        page_1_info['table_data']['category_rows'] = preview_result['category_rows']

                    apply_dict = preview_result.get('apply_dictionary', True)

                    # Step 3B: Collect remaining pages (Instant Hand-off)
                    self.update_status_safe(
                        message=f"[3/4] Finalizing remaining pages (instant hand-off)...",
                        progress=0.85,
                        detail="Assembling pre-fetched multi-core pages...",
                        text_color="#2563EB"
                    )

                    if pool_executor and prefetch_futures:
                        for fut, p_idx in prefetch_futures.items():
                            if p_idx not in prefetch_results:
                                try:
                                    res = fut.result()
                                    prefetch_results[p_idx] = res
                                except Exception as e:
                                    print(f"Error awaiting page {p_idx+1}: {e}")
                        pool_executor.shutdown(wait=False)

                        for p_idx in sorted(prefetch_results.keys()):
                            res = prefetch_results[p_idx]
                            page_num = p_idx + 1
                            t_data = res['table_data']
                            extracted_pages_data.append({
                                'page_num': page_num,
                                'table_data': t_data,
                                'ocr_results': res.get('ocr_results', [])
                            })
                            total_blocks += res.get('block_count', 0)
                            total_chars += res.get('char_count', 0)
                            row_count = len(t_data['table_rows'])
                            col_count = max((len(r) for r in t_data['table_rows']), default=1)
                            self.update_status_safe(
                                message=f"[3/4] Page {page_num}/{total_pages} processed: {row_count} rows × {col_count} cols.",
                                progress=0.85 + (0.05 * (page_num / total_pages)),
                                detail=f"Page {page_num}/{total_pages} assembled into document structure.",
                                text_color="#2563EB"
                            )
                    else:
                        # Fallback sequential loop with zero-copy
                        for idx in range(1, total_pages):
                            page_num = idx + 1
                            page_start_time = time.time()
                            base_prog = 0.28 + ((idx - 1) / max(1, total_pages - 1)) * 0.60

                            self.update_status_safe(
                                message=f"[3/4] Rendering page {page_num} of {total_pages}...",
                                progress=base_prog + (0.05 / total_pages),
                                detail=f"Page {page_num}/{total_pages} • Zero-copy rasterization...",
                                text_color="#2563EB"
                            )
                            page = doc[idx]
                            pix = page.get_pixmap(dpi=200)
                            raw = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                            if pix.n == 4:
                                img_bgr = cv2.cvtColor(raw, cv2.COLOR_RGBA2BGR) if HAS_CV2 else None
                                img_ocr = cv2.cvtColor(raw, cv2.COLOR_RGBA2RGB) if HAS_CV2 else raw
                            elif pix.n == 3:
                                img_bgr = cv2.cvtColor(raw, cv2.COLOR_RGB2BGR) if HAS_CV2 else None
                                img_ocr = raw
                            else:
                                img_bgr = raw
                                img_ocr = raw

                            self.update_status_safe(
                                message=f"[3/4] Running EasyOCR on page {page_num} of {total_pages}...",
                                progress=base_prog + (0.25 / total_pages),
                                detail=f"Batched neural inference in torch.inference_mode()...",
                                text_color="#2563EB"
                            )
                            if HAS_TORCH:
                                with torch.inference_mode():
                                    ocr_results = reader.readtext(img_ocr, batch_size=8, detail=1, paragraph=False)
                            else:
                                ocr_results = reader.readtext(img_ocr, batch_size=8, detail=1, paragraph=False)

                            self.update_status_safe(
                                message=f"[3/4] Extracting structured table grid on page {page_num} of {total_pages}...",
                                progress=base_prog + (0.45 / total_pages),
                                detail=f"Mapping column & row boundaries via spatial clustering...",
                                text_color="#2563EB"
                            )
                            table_data = StructuredTableExtractor.extract_table_from_page(
                                img_bgr, ocr_results, reader=reader, apply_dictionary=apply_dict
                            )
                            extracted_pages_data.append({
                                'page_num': page_num,
                                'table_data': table_data,
                                'ocr_results': ocr_results
                            })

                            page_duration = time.time() - page_start_time
                            total_blocks += len(ocr_results)
                            total_chars += sum(len(str(it[1])) for it in ocr_results)

                            row_count = len(table_data['table_rows'])
                            col_count = max((len(r) for r in table_data['table_rows']), default=1)
                            self.update_status_safe(
                                message=f"[3/4] Page {page_num}/{total_pages} mapped: {row_count} rows × {col_count} cols ({page_duration:.1f}s).",
                                progress=base_prog + (0.55 / total_pages),
                                detail=f"Extracted {len(ocr_results)} text fragments in {page_duration:.1f}s.",
                                text_color="#2563EB"
                            )

                    doc.close()

                elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"):
                    self.update_status_safe(
                        message=f"[3/4] Running EasyOCR on image: {os.path.basename(input_path)}...",
                        progress=0.30,
                        detail="Batched neural inference with EasyOCR...",
                        text_color="#2563EB"
                    )
                    img_bgr = cv2.imread(input_path) if HAS_CV2 else None
                    if HAS_TORCH:
                        with torch.inference_mode():
                            ocr_results = reader.readtext(input_path, batch_size=8, detail=1, paragraph=False)
                    else:
                        ocr_results = reader.readtext(input_path, batch_size=8, detail=1, paragraph=False)
                    table_data = StructuredTableExtractor.extract_table_from_page(img_bgr, ocr_results, reader=reader)

                    page_1_info = {
                        'page_num': 1,
                        'table_data': table_data,
                        'ocr_results': ocr_results
                    }
                    extracted_pages_data.append(page_1_info)
                    total_blocks = len(ocr_results)
                    total_chars = sum(len(str(it[1])) for it in ocr_results)

                    # Trigger interactive preview
                    preview_event = threading.Event()
                    preview_result = {'accepted': False}

                    def _on_preview_choice(choice):
                        if isinstance(choice, dict):
                            preview_result.update(choice)
                        else:
                            preview_result['accepted'] = bool(choice)
                        preview_event.set()

                    self.update_status_safe(
                        message=f"[Preview] Image processed ({len(table_data['table_rows'])} rows). Waiting for review...",
                        progress=0.45,
                        detail="Please review table layout and click 'Accept & Export' or 'Cancel'.",
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

                else:
                    # Fallback for office documents (e.g. pptx, html, etc.)
                    self.update_status_safe(
                        message=f"[3/4] Extracting content from: {os.path.basename(input_path)}...",
                        progress=0.50,
                        detail="Converting document text...",
                        text_color="#2563EB"
                    )
                    if HAS_MARKITDOWN:
                        md_res = self.markitdown_converter.convert(input_path)
                        text_content = md_res.text_content or ""
                    else:
                        with open(input_path, "r", encoding="utf-8", errors="replace") as f:
                            text_content = f.read()

                    lines = [line.strip() for line in text_content.splitlines() if line.strip()]
                    table_rows = [[l] for l in lines]
                    table_data = {
                        'title_headers': [],
                        'table_rows': table_rows,
                        'category_rows': set(),
                        'footer_notes': [],
                        'is_rtl': any(StructuredTableExtractor.is_arabic(l) for l in lines),
                        'has_grid': False
                    }
                    total_blocks = len(lines)
                    total_chars = sum(len(l) for l in lines)
                    extracted_pages_data.append({
                        'page_num': 1,
                        'table_data': table_data,
                        'ocr_results': []
                    })

                # Step 4: Formatting and saving to destination
                config = FORMAT_CONFIG.get(target_format, FORMAT_CONFIG["Text (.txt)"])
                self.update_status_safe(
                    message=f"[4/4] Formatting structured {config['name']}...",
                    progress=0.92,
                    detail=f"Writing output to {os.path.basename(output_path)}...",
                    text_color="#2563EB"
                )

                out_dir = os.path.dirname(os.path.abspath(output_path))
                if out_dir and not os.path.exists(out_dir):
                    os.makedirs(out_dir, exist_ok=True)

                target_ext = config["ext"]
                if target_ext == ".xlsx":
                    self._save_excel(extracted_pages_data, output_path, input_path)
                elif target_ext == ".docx":
                    self._save_word(extracted_pages_data, output_path, input_path)
                else:
                    self._save_text(extracted_pages_data, output_path, input_path)

                elapsed = time.time() - start_time
                self.after(0, self._on_conversion_success, output_path, elapsed, total_blocks, total_chars, len(extracted_pages_data))

            except Exception as e:
                self.after(0, self._on_conversion_error, str(e))

        def _on_conversion_success(self, output_path, elapsed, total_blocks, total_chars, page_count):
            self.last_converted_path = output_path
            self.set_ui_state_processing(False)

            file_size_kb = os.path.getsize(output_path) / 1024 if os.path.exists(output_path) else 0
            self.update_status_safe(
                message=f"✓ Conversion completed successfully in {elapsed:.1f}s!",
                progress=1.0,
                detail=f"Saved: {os.path.basename(output_path)} ({file_size_kb:.1f} KB) • {page_count} page(s) • Structured table columns preserved",
                text_color="#16A34A"
            )
            self.show_action_buttons()

        def _on_conversion_error(self, error_message):
            self.set_ui_state_processing(False)
            self.update_status_safe(
                message=f"✗ Conversion failed: {error_message[:60]}...",
                progress=0.0,
                detail="Please check that document is valid and dependencies are installed.",
                text_color="#DC2626"
            )
            detail = ""
            if "permission denied" in error_message.lower() or "errno 13" in error_message.lower():
                detail = "\n\nTip: The output file may currently be open in another application (such as Microsoft Excel or Word). Please close it and try again."
            messagebox.showerror("PDF2text - Conversion Error", f"Failed to convert document:\n\n{error_message}{detail}")

        # ----------------------------------------------------------------------
        # Post-Conversion Quick Actions
        # ----------------------------------------------------------------------
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
                    self.detail_label.configure(text="✓ Content / path copied to clipboard!")
                except Exception as e:
                    messagebox.showerror("PDF2text", f"Could not copy text: {e}")


# ==============================================================================
# Standard Tkinter / ttk Fallback Implementation
# ==============================================================================
else:
    class PDF2textApp(tk.Tk, ConversionEngineMixin):
        def __init__(self):
            tk.Tk.__init__(self)
            ConversionEngineMixin.__init__(self)

            self.title("PDF2text - Document & OCR Converter")
            self.geometry("620x650")
            self.minsize(560, 580)
            self.configure(bg="#F1F5F9")

            self._setup_styles()
            self._create_widgets()
            self._center_window()

        def _setup_styles(self):
            self.style = ttk.Style(self)
            try:
                self.style.theme_use("clam")
            except Exception:
                pass

            self.style.configure(".", font=("Segoe UI", 10))
            self.style.configure("Card.TFrame", background="#FFFFFF")
            self.style.configure("Header.TLabel", background="#FFFFFF", foreground="#0F172A", font=("Segoe UI", 18, "bold"))
            self.style.configure("Sub.TLabel", background="#FFFFFF", foreground="#64748B", font=("Segoe UI", 10))
            self.style.configure("FieldLabel.TLabel", background="#FFFFFF", foreground="#1E293B", font=("Segoe UI", 10, "bold"))
            self.style.configure("Status.TLabel", background="#FFFFFF", foreground="#334155", font=("Segoe UI", 10, "bold"))
            self.style.configure("Detail.TLabel", background="#FFFFFF", foreground="#64748B", font=("Segoe UI", 9))
            self.style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), background="#2563EB", foreground="#FFFFFF")
            self.style.map("Primary.TButton", background=[("active", "#1D4ED8"), ("disabled", "#93C5FD")])
            self.style.configure("Secondary.TButton", font=("Segoe UI", 10), background="#F1F5F9", foreground="#1E293B")
            self.style.map("Secondary.TButton", background=[("active", "#E2E8F0")])

            self.style.configure(
                "Format.TCombobox",
                font=("Segoe UI", 10),
                background="#FFFFFF",
                fieldbackground="#FFFFFF",
                foreground="#0F172A",
                padding=(6, 4)
            )

        def _center_window(self):
            self.update_idletasks()
            width = self.winfo_width()
            height = self.winfo_height()
            x = (self.winfo_screenwidth() // 2) - (width // 2)
            y = (self.winfo_screenheight() // 2) - (height // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")

        def _create_widgets(self):
            card = ttk.Frame(self, style="Card.TFrame", padding=24)
            card.pack(fill="both", expand=True, padx=20, pady=20)

            title_lbl = ttk.Label(card, text="PDF2text", style="Header.TLabel")
            title_lbl.pack(anchor="w")

            sub_lbl = ttk.Label(card, text="Convert scanned PDF documents and images to structured Excel, Word, or Text", style="Sub.TLabel")
            sub_lbl.pack(anchor="w", pady=(2, 14))

            ttk.Separator(card, orient="horizontal").pack(fill="x", pady=(0, 14))

            # Input File Section
            ttk.Label(card, text="Input File", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
            in_frame = ttk.Frame(card, style="Card.TFrame")
            in_frame.pack(fill="x", pady=(0, 14))

            self.input_var = tk.StringVar()
            self.input_entry = ttk.Entry(in_frame, textvariable=self.input_var, font=("Segoe UI", 10))
            self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=3)

            self.in_browse_btn = ttk.Button(in_frame, text="Browse", style="Secondary.TButton", command=self.browse_input_file)
            self.in_browse_btn.pack(side="right")

            # Output Format Section (ttk.Combobox)
            ttk.Label(card, text="Output Format", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
            self.format_combobox = ttk.Combobox(card, values=FORMAT_OPTIONS, style="Format.TCombobox", state="readonly")
            self.format_combobox.set("Excel (.xlsx)")
            self.format_combobox.pack(fill="x", pady=(0, 2), ipady=2)
            self.format_combobox.bind("<<ComboboxSelected>>", self.on_format_changed)

            self.format_desc_var = tk.StringVar(value=FORMAT_CONFIG["Excel (.xlsx)"]["desc"])
            self.format_desc_lbl = ttk.Label(card, textvariable=self.format_desc_var, style="Detail.TLabel")
            self.format_desc_lbl.pack(anchor="w", pady=(0, 14))

            # Output Destination Section
            ttk.Label(card, text="Output Destination", style="FieldLabel.TLabel").pack(anchor="w", pady=(0, 4))
            out_frame = ttk.Frame(card, style="Card.TFrame")
            out_frame.pack(fill="x", pady=(0, 16))

            self.output_var = tk.StringVar()
            self.output_entry = ttk.Entry(out_frame, textvariable=self.output_var, font=("Segoe UI", 10))
            self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=3)

            self.out_browse_btn = ttk.Button(out_frame, text="Browse", style="Secondary.TButton", command=self.browse_output_file)
            self.out_browse_btn.pack(side="right")

            # Convert Button
            self.convert_btn = ttk.Button(card, text="Run OCR & Export to Excel (.xlsx)", style="Primary.TButton", command=self.start_conversion)
            self.convert_btn.pack(fill="x", ipady=6, pady=(0, 10))

            # Progress Bar
            self.progress_bar = ttk.Progressbar(card, orient="horizontal", mode="determinate", maximum=100)
            self.progress_bar.pack(fill="x", pady=(0, 8))

            # Status & Details
            self.status_var = tk.StringVar(value="Ready. Select a document and click Convert.")
            self.status_label = ttk.Label(card, textvariable=self.status_var, style="Status.TLabel")
            self.status_label.pack(anchor="w")

            self.detail_var = tk.StringVar(value="Step: Idle • Engine: EasyOCR & OpenCV Morphology Table Extraction")
            self.detail_label = ttk.Label(card, textvariable=self.detail_var, style="Detail.TLabel")
            self.detail_label.pack(anchor="w", pady=(2, 0))

            # Actions Row
            self.actions_frame = ttk.Frame(card, style="Card.TFrame")
            self.actions_frame.pack(fill="x", pady=(12, 0))

            self.open_file_btn = ttk.Button(self.actions_frame, text="📄 Open File", style="Secondary.TButton", command=self.open_converted_file)
            self.open_folder_btn = ttk.Button(self.actions_frame, text="📁 Open Folder", style="Secondary.TButton", command=self.open_containing_folder)
            self.copy_btn = ttk.Button(self.actions_frame, text="📋 Copy Text", style="Secondary.TButton", command=self.copy_to_clipboard)

        def get_input_path(self):
            return self.input_var.get().strip()

        def set_input_path(self, path):
            self.input_var.set(path)

        def get_output_path(self):
            return self.output_var.get().strip()

        def set_output_path(self, path):
            self.output_var.set(path)

        def set_ui_state_processing(self, is_processing: bool):
            if is_processing:
                self.convert_btn.configure(state="disabled", text="Processing...")
                self.in_browse_btn.configure(state="disabled")
                self.out_browse_btn.configure(state="disabled")
                self.format_combobox.configure(state="disabled")
            else:
                fmt = self.format_combobox.get().strip()
                self.convert_btn.configure(state="normal", text=f"Run OCR & Export to {fmt}")
                self.in_browse_btn.configure(state="normal")
                self.out_browse_btn.configure(state="normal")
                self.format_combobox.configure(state="readonly")

        def hide_action_buttons(self):
            self.open_file_btn.pack_forget()
            self.open_folder_btn.pack_forget()
            self.copy_btn.pack_forget()

        def show_action_buttons(self):
            self.open_file_btn.pack(side="left", padx=(0, 8))
            self.open_folder_btn.pack(side="left", padx=(0, 8))
            self.copy_btn.pack(side="left")

        def update_status_safe(self, message: str, progress: float = None, detail: str = None, text_color: str = None):
            def _apply():
                if message:
                    self.status_var.set(message)
                if detail:
                    self.detail_var.set(detail)
                if progress is not None:
                    clamped = max(0.0, min(1.0, float(progress)))
                    self.progress_bar["value"] = clamped * 100
                self.update_idletasks()

            self.after(0, _apply)

        def on_format_changed(self, event=None):
            fmt = self.format_combobox.get().strip()
            config = FORMAT_CONFIG.get(fmt, FORMAT_CONFIG["Text (.txt)"])
            target_ext = config["ext"]

            self.convert_btn.configure(text=f"Run OCR & Export to {fmt}")
            self.format_desc_var.set(config["desc"])

            curr_out = self.get_output_path()
            if curr_out:
                base, _ = os.path.splitext(curr_out)
                self.set_output_path(base + target_ext)

        def browse_input_file(self):
            selected = filedialog.askopenfilename(title="PDF2text - Select Input Document", filetypes=FILE_TYPES)
            if selected:
                clean_path = os.path.normpath(selected)
                self.set_input_path(clean_path)

                fmt = self.format_combobox.get().strip()
                config = FORMAT_CONFIG.get(fmt, FORMAT_CONFIG["Text (.txt)"])
                suggested_out = os.path.splitext(clean_path)[0] + config["ext"]
                self.set_output_path(suggested_out)

                self.update_status_safe(
                    message=f"Document loaded. Ready to convert to {fmt}.",
                    progress=0.0,
                    detail=f"Source: {os.path.basename(clean_path)}"
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
            self.update_status_safe(
                message="Operation cancelled by user.",
                progress=0.0,
                detail="No files were exported. Ready for next document."
            )

        def _run_conversion_worker(self, input_path: str, output_path: str, target_format: str):
            start_time = time.time()
            try:
                self.update_status_safe(
                    message="[1/4] Preparing OCR engine (Arabic & English)...",
                    progress=0.05,
                    detail="Loading cached EasyOCR weights..."
                )
                reader = self._get_ocr_reader()

                ext = os.path.splitext(input_path)[1].lower()
                self.update_status_safe(
                    message=f"[2/4] Opening document: {os.path.basename(input_path)}...",
                    progress=0.12,
                    detail="Analyzing document structure..."
                )

                extracted_pages_data = []
                total_blocks = 0
                total_chars = 0

                if ext == ".pdf":
                    if not HAS_PYMUPDF:
                        raise RuntimeError("PyMuPDF is required to process PDF documents. Run 'pip install pymupdf'.")

                    doc = pymupdf.open(input_path)
                    total_pages = len(doc)
                    self.update_status_safe(
                        message=f"[2/4] Document loaded: {total_pages} page(s) detected.",
                        progress=0.15,
                        detail="Rendering Page 1 for interactive preview & verification..."
                    )

                    # Step 3A: Process Page 1 with Zero-Copy Rasterization & Batched Inference
                    self.update_status_safe(
                        message=f"[3/4] Rendering Page 1 of {total_pages}...",
                        progress=0.18,
                        detail="Zero-copy rasterization at 200 DPI..."
                    )
                    page0 = doc[0]
                    pix0 = page0.get_pixmap(dpi=200)
                    raw0 = np.frombuffer(pix0.samples, dtype=np.uint8).reshape(pix0.h, pix0.w, pix0.n)
                    if pix0.n == 4:
                        img_bgr0 = cv2.cvtColor(raw0, cv2.COLOR_RGBA2BGR) if HAS_CV2 else None
                        img_ocr0 = cv2.cvtColor(raw0, cv2.COLOR_RGBA2RGB) if HAS_CV2 else raw0
                    elif pix0.n == 3:
                        img_bgr0 = cv2.cvtColor(raw0, cv2.COLOR_RGB2BGR) if HAS_CV2 else None
                        img_ocr0 = raw0
                    else:
                        img_bgr0 = raw0
                        img_ocr0 = raw0

                    self.update_status_safe(
                        message=f"[3/4] Running EasyOCR on Page 1 of {total_pages}...",
                        progress=0.22,
                        detail="Batched neural inference in torch.inference_mode()..."
                    )
                    if HAS_TORCH:
                        with torch.inference_mode():
                            ocr_res0 = reader.readtext(img_ocr0, batch_size=8, detail=1, paragraph=False)
                    else:
                        ocr_res0 = reader.readtext(img_ocr0, batch_size=8, detail=1, paragraph=False)

                    self.update_status_safe(
                        message=f"[3/4] Extracting structured table grid on Page 1...",
                        progress=0.25,
                        detail="Applying spatial clustering, RTL alignment & category merging..."
                    )
                    table_data0 = StructuredTableExtractor.extract_table_from_page(img_bgr0, ocr_res0, reader=reader, apply_dictionary=True)
                    page_1_info = {
                        'page_num': 1,
                        'table_data': table_data0,
                        'ocr_results': ocr_res0
                    }
                    extracted_pages_data.append(page_1_info)
                    total_blocks += len(ocr_res0)
                    total_chars += sum(len(str(it[1])) for it in ocr_res0)

                    # Launch Background Prefetch Pipeline for Pages 2..N
                    pool_executor = None
                    prefetch_futures = {}
                    prefetch_results = {}
                    prefetch_cancelled = threading.Event()
                    num_workers = min(4, max(1, os.cpu_count() // 2))

                    if total_pages > 1:
                        models_dir = self._get_model_storage_dir()
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
                                            detail=f"Previewing Page 1 • Prefetched {completed_count}/{total_pages - 1} background pages..."
                                        )
                                    except Exception as e:
                                        print(f"Prefetch page {p_idx+1} error: {e}")

                            threading.Thread(target=_monitor_prefetch, daemon=True).start()
                        except Exception as pool_err:
                            print(f"Notice: ProcessPoolExecutor fallback: {pool_err}")
                            pool_executor = None

                    # Trigger interactive preview modal
                    preview_event = threading.Event()
                    preview_result = {'accepted': False}

                    def _on_preview_choice(choice):
                        if isinstance(choice, dict):
                            preview_result.update(choice)
                        else:
                            preview_result['accepted'] = bool(choice)
                        preview_event.set()

                    self.update_status_safe(
                        message=f"[Preview] Page 1 ready ({len(table_data0['table_rows'])} rows). Review while remaining pages prefetch...",
                        progress=0.28,
                        detail=f"Parallel pipeline active across {num_workers} CPU cores. Review Page 1 table layout."
                    )
                    self.after(0, self._open_page_1_preview, page_1_info, total_pages, target_format, _on_preview_choice)

                    # Block worker thread until user decides in the preview dialog
                    preview_event.wait()

                    if not preview_result.get('accepted', False):
                        prefetch_cancelled.set()
                        if pool_executor:
                            pool_executor.shutdown(wait=False, cancel_futures=True)
                        doc.close()
                        self.after(0, self._on_conversion_cancelled)
                        return

                    # Update page 1 with any user edits made in preview
                    if 'table_rows' in preview_result and preview_result['table_rows']:
                        page_1_info['table_data']['table_rows'] = preview_result['table_rows']
                    if 'category_rows' in preview_result and preview_result['category_rows'] is not None:
                        page_1_info['table_data']['category_rows'] = preview_result['category_rows']

                    apply_dict = preview_result.get('apply_dictionary', True)

                    # Step 3B: Collect remaining pages (Instant Hand-off)
                    self.update_status_safe(
                        message=f"[3/4] Finalizing remaining pages (instant hand-off)...",
                        progress=0.85,
                        detail="Assembling pre-fetched multi-core pages..."
                    )

                    if pool_executor and prefetch_futures:
                        for fut, p_idx in prefetch_futures.items():
                            if p_idx not in prefetch_results:
                                try:
                                    res = fut.result()
                                    prefetch_results[p_idx] = res
                                except Exception as e:
                                    print(f"Error awaiting page {p_idx+1}: {e}")
                        pool_executor.shutdown(wait=False)

                        for p_idx in sorted(prefetch_results.keys()):
                            res = prefetch_results[p_idx]
                            page_num = p_idx + 1
                            t_data = res['table_data']
                            extracted_pages_data.append({
                                'page_num': page_num,
                                'table_data': t_data,
                                'ocr_results': res.get('ocr_results', [])
                            })
                            total_blocks += res.get('block_count', 0)
                            total_chars += res.get('char_count', 0)
                            row_count = len(t_data['table_rows'])
                            col_count = max((len(r) for r in t_data['table_rows']), default=1)
                            self.update_status_safe(
                                message=f"[3/4] Page {page_num}/{total_pages} processed: {row_count} rows × {col_count} cols.",
                                progress=0.85 + (0.05 * (page_num / total_pages)),
                                detail=f"Page {page_num}/{total_pages} assembled into document structure."
                            )
                    else:
                        # Fallback sequential loop with zero-copy
                        for idx in range(1, total_pages):
                            page_num = idx + 1
                            page_start_time = time.time()
                            base_prog = 0.28 + ((idx - 1) / max(1, total_pages - 1)) * 0.60

                            self.update_status_safe(
                                message=f"[3/4] Rendering page {page_num} of {total_pages}...",
                                progress=base_prog + (0.05 / total_pages),
                                detail=f"Page {page_num}/{total_pages} • Zero-copy rasterization..."
                            )

                            page = doc[idx]
                            pix = page.get_pixmap(dpi=200)
                            raw = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                            if pix.n == 4:
                                img_bgr = cv2.cvtColor(raw, cv2.COLOR_RGBA2BGR) if HAS_CV2 else None
                                img_ocr = cv2.cvtColor(raw, cv2.COLOR_RGBA2RGB) if HAS_CV2 else raw
                            elif pix.n == 3:
                                img_bgr = cv2.cvtColor(raw, cv2.COLOR_RGB2BGR) if HAS_CV2 else None
                                img_ocr = raw
                            else:
                                img_bgr = raw
                                img_ocr = raw

                            self.update_status_safe(
                                message=f"[3/4] Running EasyOCR on page {page_num} of {total_pages}...",
                                progress=base_prog + (0.25 / total_pages),
                                detail=f"Batched neural inference in torch.inference_mode()..."
                            )

                            if HAS_TORCH:
                                with torch.inference_mode():
                                    ocr_results = reader.readtext(img_ocr, batch_size=8, detail=1, paragraph=False)
                            else:
                                ocr_results = reader.readtext(img_ocr, batch_size=8, detail=1, paragraph=False)

                            self.update_status_safe(
                                message=f"[3/4] Extracting structured table grid on page {page_num} of {total_pages}...",
                                progress=base_prog + (0.45 / total_pages),
                                detail=f"Mapping column & row boundaries via spatial clustering..."
                            )

                            table_data = StructuredTableExtractor.extract_table_from_page(
                                img_bgr, ocr_results, reader=reader, apply_dictionary=apply_dict
                            )

                            extracted_pages_data.append({
                                'page_num': page_num,
                                'table_data': table_data,
                                'ocr_results': ocr_results
                            })

                            page_duration = time.time() - page_start_time
                            total_blocks += len(ocr_results)
                            total_chars += sum(len(str(it[1])) for it in ocr_results)

                            row_count = len(table_data['table_rows'])
                            col_count = max((len(r) for r in table_data['table_rows']), default=1)
                            grid_type = "Grid Table" if table_data['has_grid'] else "Spatial Table"

                            self.update_status_safe(
                                message=f"[3/4] Page {page_num}/{total_pages} mapped: {row_count} rows × {col_count} cols ({grid_type}, {page_duration:.1f}s).",
                                progress=base_prog + (0.60 / total_pages),
                                detail=f"Extracted {total_blocks:,} blocks across {page_num}/{total_pages} pages..."
                            )

                    doc.close()

                elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"):
                    self.update_status_safe(
                        message=f"[3/4] Running EasyOCR on image: {os.path.basename(input_path)}...",
                        progress=0.30,
                        detail="Batched neural inference with EasyOCR..."
                    )
                    img_bgr = cv2.imread(input_path) if HAS_CV2 else None
                    if HAS_TORCH:
                        with torch.inference_mode():
                            ocr_results = reader.readtext(input_path, batch_size=8, detail=1, paragraph=False)
                    else:
                        ocr_results = reader.readtext(input_path, batch_size=8, detail=1, paragraph=False)
                    table_data = StructuredTableExtractor.extract_table_from_page(img_bgr, ocr_results, reader=reader)
                    table_data = StructuredTableExtractor.extract_table_from_page(img_bgr, ocr_results, reader=reader)

                    page_1_info = {
                        'page_num': 1,
                        'table_data': table_data,
                        'ocr_results': ocr_results
                    }
                    extracted_pages_data.append(page_1_info)
                    total_blocks = len(ocr_results)
                    total_chars = sum(len(str(it[1])) for it in ocr_results)

                    # Trigger interactive preview
                    preview_event = threading.Event()
                    preview_result = {'accepted': False}

                    def _on_preview_choice(choice):
                        if isinstance(choice, dict):
                            preview_result.update(choice)
                        else:
                            preview_result['accepted'] = bool(choice)
                        preview_event.set()

                    self.update_status_safe(
                        message=f"[Preview] Image processed ({len(table_data['table_rows'])} rows). Waiting for review...",
                        progress=0.45,
                        detail="Please review table layout and click 'Accept & Export' or 'Cancel'."
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

                else:
                    self.update_status_safe(
                        message=f"[3/4] Extracting content from: {os.path.basename(input_path)}...",
                        progress=0.50,
                        detail="Converting document text..."
                    )
                    if HAS_MARKITDOWN:
                        md_res = self.markitdown_converter.convert(input_path)
                        text_content = md_res.text_content or ""
                    else:
                        with open(input_path, "r", encoding="utf-8", errors="replace") as f:
                            text_content = f.read()

                    lines = [line.strip() for line in text_content.splitlines() if line.strip()]
                    table_rows = [[l] for l in lines]
                    table_data = {
                        'title_headers': [],
                        'table_rows': table_rows,
                        'category_rows': set(),
                        'footer_notes': [],
                        'is_rtl': any(StructuredTableExtractor.is_arabic(l) for l in lines),
                        'has_grid': False
                    }
                    total_blocks = len(lines)
                    total_chars = sum(len(l) for l in lines)
                    extracted_pages_data.append({
                        'page_num': 1,
                        'table_data': table_data,
                        'ocr_results': []
                    })

                config = FORMAT_CONFIG.get(target_format, FORMAT_CONFIG["Text (.txt)"])
                self.update_status_safe(
                    message=f"[4/4] Formatting structured {config['name']}...",
                    progress=0.92,
                    detail=f"Writing output to {os.path.basename(output_path)}..."
                )

                out_dir = os.path.dirname(os.path.abspath(output_path))
                if out_dir and not os.path.exists(out_dir):
                    os.makedirs(out_dir, exist_ok=True)

                target_ext = config["ext"]
                if target_ext == ".xlsx":
                    self._save_excel(extracted_pages_data, output_path, input_path)
                elif target_ext == ".docx":
                    self._save_word(extracted_pages_data, output_path, input_path)
                else:
                    self._save_text(extracted_pages_data, output_path, input_path)

                elapsed = time.time() - start_time
                self.after(0, self._on_conversion_success, output_path, elapsed, total_blocks, total_chars, len(extracted_pages_data))

            except Exception as e:
                self.after(0, self._on_conversion_error, str(e))

        def _on_conversion_success(self, output_path, elapsed, total_blocks, total_chars, page_count):
            self.last_converted_path = output_path
            self.set_ui_state_processing(False)

            file_size_kb = os.path.getsize(output_path) / 1024 if os.path.exists(output_path) else 0
            self.update_status_safe(
                message=f"✓ Conversion completed successfully in {elapsed:.1f}s!",
                progress=1.0,
                detail=f"Saved: {os.path.basename(output_path)} ({file_size_kb:.1f} KB) • {page_count} page(s) • Structured table columns preserved"
            )
            self.show_action_buttons()

        def _on_conversion_error(self, error_message):
            self.set_ui_state_processing(False)
            self.update_status_safe(
                message=f"✗ Conversion failed: {error_message[:60]}...",
                progress=0.0,
                detail="Please check that document is valid and dependencies are installed."
            )
            detail = ""
            if "permission denied" in error_message.lower() or "errno 13" in error_message.lower():
                detail = "\n\nTip: The output file may currently be open in another application (such as Microsoft Excel or Word). Please close it and try again."
            messagebox.showerror("PDF2text - Conversion Error", f"Failed to convert document:\n\n{error_message}{detail}")

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
                    self.detail_var.set("✓ Content / path copied to clipboard!")
                except Exception as e:
                    messagebox.showerror("PDF2text", f"Could not copy text: {e}")


def main():
    app = PDF2textApp()
    app.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
