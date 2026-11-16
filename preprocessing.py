pip install pymupdf python-docx pytesseract pillow

# Preprocessing.py

import re
import os
from pathlib import Path

import fitz  # PyMuPDF
import pytesseract

from PIL import Image
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table


# ------------------------------------------
# CONFIGURATION
# ------------------------------------------

SUPPORTED_FORMATS = {".pdf", ".docx", ".txt"}

MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB


# ------------------------------------------
# 1. TEXT NORMALIZATION
# ------------------------------------------

def normalize_text(text):
    """
    Clean and normalize extracted text while
    preserving sentence structure and paragraph breaks.
    """

    # Normalize line endings
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    # Remove null characters
    text = text.replace("\x00", "")

    # Replace tabs with spaces
    text = text.replace("\t", " ")

    # Remove excessive spaces
    text = re.sub(r"[ ]{2,}", " ", text)

    # Remove excessive blank lines
    text = re.sub(r"\n[ ]*\n+", "\n\n", text)

    # Remove leading and trailing whitespace
    text = text.strip()

    return text


# ------------------------------------------
# 2. PDF TEXT EXTRACTION
# ------------------------------------------

def extract_pdf(file_path):
    """
    Extract text from PDF documents.

    If a page contains little or no extractable text,
    OCR is used to process that page.
    """

    pages = []

    pdf = fitz.open(file_path)

    try:
        for page_number, page in enumerate(pdf, start=1):

            # Attempt normal PDF text extraction
            text = page.get_text("text")

            method = "text_extraction"

            # Use OCR when the page is likely scanned
            if len(text.strip()) < 20:

                pix = page.get_pixmap(
                    matrix=fitz.Matrix(2, 2),
                    alpha=False
                )

                image = Image.frombytes(
                    "RGB",
                    [pix.width, pix.height],
                    pix.samples
                )

                try:
                    text = pytesseract.image_to_string(
                        image,
                        lang="eng"
                    )
                except pytesseract.TesseractNotFoundError:
                    raise RuntimeError(
                        "Tesseract OCR is not installed or "
                        "cannot be found on this system."
                    )

                method = "OCR"

            pages.append({
                "page_number": page_number,
                "text": normalize_text(text),
                "extraction_method": method
            })

    finally:
        pdf.close()

    return pages


# ------------------------------------------
# 3. WORD DOCUMENT EXTRACTION
# ------------------------------------------

def extract_docx(file_path):
    """
    Extract paragraphs and tables from Word documents.
    Maintains the order of paragraphs and tables.
    """

    document = Document(file_path)

    extracted_blocks = []

    for block in document.element.body:

        if block.tag.endswith("}p"):

            paragraph = Paragraph(
                block,
                document
            )

            if paragraph.text.strip():
                extracted_blocks.append(paragraph.text)

        elif block.tag.endswith("}tbl"):

            table = Table(
                block,
                document
            )

            for row in table.rows:

                row_text = [
                    cell.text.strip()
                    for cell in row.cells
                ]

                extracted_blocks.append(
                    " | ".join(row_text)
                )

    full_text = "\n\n".join(extracted_blocks)

    return [{
        "page_number": None,
        "text": normalize_text(full_text),
        "extraction_method": "docx_extraction"
    }]


# ------------------------------------------
# 4. PLAIN TEXT EXTRACTION
# ------------------------------------------

def extract_txt(file_path):
    """
    Extract text from TXT files.
    """

    try:
        with open(
            file_path,
            "r",
            encoding="utf-8-sig"
        ) as file:

            text = file.read()

    except UnicodeDecodeError:

        # Fallback for older Windows text files
        with open(
            file_path,
            "r",
            encoding="cp1252"
        ) as file:

            text = file.read()

    return [{
        "page_number": None,
        "text": normalize_text(text),
        "extraction_method": "txt_extraction"
    }]


# ------------------------------------------
# 5. MAIN DOCUMENT PREPROCESSING FUNCTION
# ------------------------------------------

def preprocess_document(file_path):
    """
    Main entry point for document preprocessing.

    Accepts PDF, DOCX, and TXT files.

    Returns extracted and normalized text,
    document metadata, and extraction details.
    """

    file_path = Path(file_path)

    # Verify document exists
    if not file_path.exists():
        raise FileNotFoundError(
            f"Document not found: {file_path}"
        )

    # Verify supported file format
    extension = file_path.suffix.lower()

    if extension not in SUPPORTED_FORMATS:
        raise ValueError(
            f"Unsupported document format: {extension}"
        )

    # Verify file size
    if file_path.stat().st_size > MAX_FILE_SIZE:
        raise ValueError(
            "Document exceeds the 50 MB size limit."
        )

    # Select extraction method
    if extension == ".pdf":
        pages = extract_pdf(file_path)

    elif extension == ".docx":
        pages = extract_docx(file_path)

    elif extension == ".txt":
        pages = extract_txt(file_path)

    # Combine extracted pages
    full_text = "\n\n".join(
        page["text"]
        for page in pages
        if page["text"]
    )

    full_text = normalize_text(full_text)

    if not full_text:
        raise ValueError(
            "No readable text was extracted from the document."
        )

    # Document processing results
    result = {
        "filename": file_path.name,
        "file_type": extension,
        "page_count": len(pages),
        "character_count": len(full_text),
        "text": full_text,
        "pages": pages
    }

    return result


# ------------------------------------------
# 6. TEST THE PREPROCESSING MODULE
# ------------------------------------------

if __name__ == "__main__":

    document_path = "test_report.pdf"

    try:

        result = preprocess_document(document_path)

        print("\nDOCUMENT PROCESSING COMPLETE")

        print("Filename:", result["filename"])
        print("File Type:", result["file_type"])
        print("Pages:", result["page_count"])
        print("Characters:", result["character_count"])

        print("\nEXTRACTED NARRATIVE:")
        print(result["text"])

    except Exception as error:

        print("Processing error:", error)












## Beginning of AI Pipeline

from preprocessing import preprocess_document
from pii_detection import detect_pii

# Step 1: Process the police report
document = preprocess_document("test_report.pdf")

# Step 2: Retrieve normalized narrative
narrative = document["text"]

# Step 3: Identify PII
detected_pii = detect_pii(narrative)

# Step 4: Display results
for item in detected_pii:
    print(item)
