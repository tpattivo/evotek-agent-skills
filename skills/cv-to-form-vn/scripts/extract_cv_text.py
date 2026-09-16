"""
Extract text from a CV PDF.
Usage: python extract_cv_text.py <path-to-pdf>
"""
import sys, pdfplumber

if len(sys.argv) < 2:
    print('Usage: python extract_cv_text.py <path-to-pdf>')
    sys.exit(1)

path = sys.argv[1]
with pdfplumber.open(path) as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        print(f'=== PAGE {i+1} ===')
        print(text or '(no text extracted)')
