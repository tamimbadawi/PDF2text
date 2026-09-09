"""
Headless OCR verification script for Arabic scanned PDF documents.
Tests PyMuPDF (fitz) rendering + EasyOCR Arabic/English text extraction.
Outputs clean UTF-8 text file.
"""

import os
import sys
import time
import warnings
from pathlib import Path

# Suppress PyTorch deprecation and device warnings for clean output
warnings.filterwarnings("ignore", category=UserWarning)

# Ensure console supports UTF-8 on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import pymupdf  # Modern PyMuPDF import
import easyocr


def run_ocr_pipeline(pdf_path: str, output_txt_path: str, dpi: int = 150, max_pages: int = None):
    print("=" * 60)
    print("PDF to Text - EasyOCR Headless Verification")
    print("=" * 60)

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Input PDF not found: {pdf_path}")

    start_total = time.time()

    print(f"\n[1/4] Initializing EasyOCR reader (Arabic + English, CPU mode)...")
    init_start = time.time()
    reader = easyocr.Reader(["ar", "en"], gpu=False, verbose=False)
    print(f"      Reader initialized in {time.time() - init_start:.2f}s")

    print(f"\n[2/4] Opening document: {pdf_path}")
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    pages_to_process = min(total_pages, max_pages) if max_pages else total_pages
    print(f"      Total document pages: {total_pages}")
    print(f"      Pages to process: {pages_to_process} (Render DPI: {dpi})")

    extracted_pages = []
    total_characters = 0

    print(f"\n[3/4] Processing pages...")
    for idx in range(pages_to_process):
        page_start = time.time()
        page_num = idx + 1
        page = doc[idx]

        # In-memory rendering (no disk temp files)
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")

        # Perform EasyOCR extraction
        lines = reader.readtext(img_bytes, detail=0, paragraph=False)
        page_text = "\n".join(line.strip() for line in lines if line.strip())

        page_duration = time.time() - page_start
        char_count = len(page_text)
        total_characters += char_count

        print(f"      - Page {page_num}/{pages_to_process}: {len(lines)} blocks, {char_count} chars in {page_duration:.2f}s")

        header = f"--- [ صفحة / Page {page_num} ] ---"
        extracted_pages.append(f"{header}\n\n{page_text}")

    doc.close()

    print(f"\n[4/4] Saving output to: {output_txt_path}")
    out_dir = os.path.dirname(os.path.abspath(output_txt_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    full_text = "\n\n" + ("\n\n" + "=" * 50 + "\n\n").join(extracted_pages) + "\n"
    with open(output_txt_path, "w", encoding="utf-8") as f:
        f.write(full_text)

    total_duration = time.time() - start_total
    file_size_kb = os.path.getsize(output_txt_path) / 1024

    print("=" * 60)
    print("Verification Summary:")
    print(f"  • Total pages processed: {pages_to_process}/{total_pages}")
    print(f"  • Total characters extracted: {total_characters:,}")
    print(f"  • Output file: {output_txt_path} ({file_size_kb:.1f} KB)")
    print(f"  • Total pipeline duration: {total_duration:.2f}s ({total_duration/pages_to_process:.2f}s/page avg)")
    print("=" * 60)

    return {
        "pages_processed": pages_to_process,
        "total_pages": total_pages,
        "total_characters": total_characters,
        "duration_seconds": total_duration,
        "output_path": output_txt_path
    }


if __name__ == "__main__":
    input_pdf = r"D:\Brave\اسعار السعودية يونيو 2026.pdf"
    output_txt = r"D:\Brave\اسعار السعودية يونيو 2026.txt"
    
    # Process all pages
    run_ocr_pipeline(input_pdf, output_txt, dpi=150)
