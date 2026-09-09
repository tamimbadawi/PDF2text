import fitz  # PyMuPDF to convert PDF pages to images
import easyocr

# 1. Initialize EasyOCR reader for Arabic and English
print("Loading OCR models (this may take a moment)...")
reader = easyocr.Reader(['ar', 'en'], gpu=False)  # Set gpu=True if you have an Nvidia GPU

# 2. Open your scanned PDF
pdf_path = r"D:\Brave\اسعار السعودية يونيو 2026.pdf"
doc = fitz.open(pdf_path)

full_markdown = []

# 3. Loop through every page, render to image, and extract text
for page_num in range(len(doc)):
    page = doc[page_num]
    pix = page.get_pixmap(dpi=300) # High DPI for better Arabic OCR quality
    image_path = f"temp_page_{page_num}.png"
    pix.save(image_path)
    
    print(f"Processing OCR for page {page_num + 1}...")
    results = reader.readtext(image_path, detail=0) # detail=0 returns just the text strings
    
    full_markdown.append(f"# Page {page_num + 1}\n\n" + " ".join(results))

# 4. Save results to a clean Markdown file
output_path = r"D:\Brave\output_ocr.md"
with open(output_path, "w", encoding="utf-8") as f:
    f.write("\n\n".join(full_markdown))

print(f"Done! OCR text saved successfully to: {output_path}")
