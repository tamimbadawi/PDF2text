# PDF2text - Arabic/English PDF to Excel / Word / Text Converter

A standalone Windows desktop application that converts scanned PDF documents (Arabic and English) into structured Excel (.xlsx), Word (.docx), or plain-text (.txt) files using spatial bounding-box OCR and table-layout detection.

## Features
- Spatial table extraction - bounding-box clustering maps text into correct Excel columns
- Arabic RTL support - proper right-to-left reading order with mixed numeric handling
- Merged section headers - category banners auto-merged across columns via openpyxl
- Multiprocessing - ProcessPoolExecutor processes pages in parallel while you preview page 1
- Background prefetch - remaining pages fully parsed before you click Export
- Arabic dictionary correction - fuzzy-matched against dictionary.json
- Overwrite confirmation - prompts before replacing an existing output file
- No-console GUI - launches instantly via app.pyw / PDF2text.exe

## Supported Formats
| Input | Output |
|-------|--------|
| Scanned PDF (Arabic / English) | Excel (.xlsx) |
| Image PDF | Word (.docx) |
| | Plain text (.txt) |

## Quick Start (from source)
```bash
pip install -r requirements.txt
python app.pyw
```n
## Build Standalone EXE
```bash
python -m PyInstaller --clean -y PDF2text.spec
# Output: dist/PDF2text/PDF2text.exe
```n
## Project Structure
```n app.py / app.pyw    Main application (GUI + conversion engine)
 PDF2text.spec       PyInstaller build specification
 dictionary.json     Arabic OCR correction dictionary
 requirements.txt    Python dependencies
 models/             EasyOCR model weights (downloaded on first run)
```
