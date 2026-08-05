# backend/ai_service/parsers.py
import os
import pdfplumber
import pandas as pd
from docx import Document

# Files larger than this are rejected before parsing (bytes)
MAX_FILE_SIZE = 25 * 1024 * 1024  # 25 MB

ALLOWED_DOC_EXTENSIONS = {'.pdf', '.doc', '.docx'}
ALLOWED_EXCEL_EXTENSIONS = {'.xls', '.xlsx', '.csv'}


class FileValidationError(Exception):
    """Raised when an uploaded file fails basic sanity checks."""
    pass


def validate_uploaded_file(uploaded_file, allowed_extensions, label="File"):
    """
    Basic guardrails every uploaded file should pass before any parser touches it.
    Raises FileValidationError with a clean, user-facing message on failure.
    """
    if uploaded_file is None:
        raise FileValidationError(f"{label} is missing.")

    if uploaded_file.size == 0:
        raise FileValidationError(f"{label} is empty (0 bytes). Please re-upload.")

    if uploaded_file.size > MAX_FILE_SIZE:
        size_mb = uploaded_file.size / (1024 * 1024)
        raise FileValidationError(
            f"{label} is {size_mb:.1f}MB, which exceeds the {MAX_FILE_SIZE // (1024*1024)}MB limit."
        )

    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext not in allowed_extensions:
        raise FileValidationError(
            f"{label} has an unsupported extension '{ext}'. "
            f"Allowed: {', '.join(sorted(allowed_extensions))}"
        )

    return ext


def extract_text_from_pdf(pdf_file) -> str:
    """
    Extracts all text and basic table structures from a PDF file.
    """
    text_content = []

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    text_content.append(f"--- PAGE {page_num + 1} ---\n{text}")

                tables = page.extract_tables()
                for table in tables:
                    if not table or not table[0]:
                        continue
                    df = pd.DataFrame(table[1:], columns=table[0])
                    text_content.append(f"\n[TABLE ON PAGE {page_num + 1}]\n{df.to_string(index=False)}\n")
    except Exception as e:
        raise FileValidationError(f"Could not read PDF '{pdf_file.name}': {str(e)}")

    if not text_content:
        raise FileValidationError(
            f"No extractable text found in '{pdf_file.name}'. "
            "It may be a scanned/image-only PDF that needs OCR."
        )

    return "\n".join(text_content)


def extract_text_from_docx(docx_file) -> str:
    """
    Extracts text and table content from a Word (.docx) file.
    """
    try:
        doc = Document(docx_file)
    except Exception as e:
        raise FileValidationError(f"Could not read Word document '{docx_file.name}': {str(e)}")

    text_content = []

    for para in doc.paragraphs:
        if para.text.strip():
            text_content.append(para.text)

    for table_num, table in enumerate(doc.tables):
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        if rows:
            df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
            text_content.append(f"\n[TABLE {table_num + 1}]\n{df.to_string(index=False)}\n")

    if not text_content:
        raise FileValidationError(f"No extractable text found in '{docx_file.name}'.")

    return "\n".join(text_content)


def extract_document_text(doc_file) -> str:
    """
    Routes a document (OM or T12) to the right parser based on its extension.
    Supports: .pdf, .doc, .docx
    """
    ext = validate_uploaded_file(doc_file, ALLOWED_DOC_EXTENSIONS, label=doc_file.name)

    if ext == '.pdf':
        return extract_text_from_pdf(doc_file)
    elif ext in ('.doc', '.docx'):
        return extract_text_from_docx(doc_file)

    # Should be unreachable given validate_uploaded_file, but fail loudly if it happens
    raise FileValidationError(f"Unhandled document extension '{ext}' for '{doc_file.name}'.")


def extract_t12_text(t12_file) -> str:
    """
    T12 can legitimately be a PDF/Word doc OR an Excel/CSV file — route based on extension.
    """
    ext = os.path.splitext(t12_file.name)[1].lower()

    if ext in ALLOWED_EXCEL_EXTENSIONS:
        return extract_data_from_excel(t12_file)
    elif ext in ALLOWED_DOC_EXTENSIONS:
        return extract_document_text(t12_file)
    else:
        raise FileValidationError(
            f"'{t12_file.name}' has an unsupported extension '{ext}'. "
            f"Allowed: {', '.join(sorted(ALLOWED_DOC_EXTENSIONS | ALLOWED_EXCEL_EXTENSIONS))}"
        )


def extract_data_from_excel(excel_file) -> str:
    """
    Extracts data from an Excel/CSV file and formats it as a clean string.
    Resets the file pointer to 0 before AND after reading, so callers can
    safely re-read the same file object afterward (e.g. into a DataFrame).
    """
    ext = validate_uploaded_file(excel_file, ALLOWED_EXCEL_EXTENSIONS, label=excel_file.name)

    try:
        excel_file.seek(0)
        if ext == '.csv':
            df = pd.read_csv(excel_file)
        else:
            df = pd.read_excel(excel_file, sheet_name=0)

        df = df.dropna(how='all').dropna(axis=1, how='all')

        if df.empty:
            raise FileValidationError(f"'{excel_file.name}' has no usable data after cleaning empty rows/columns.")

        result = df.to_string(index=False)
    except FileValidationError:
        raise
    except Exception as e:
        raise FileValidationError(f"Could not read '{excel_file.name}': {str(e)}")
    finally:
        # Leave the pointer reset so the caller can safely read it again
        excel_file.seek(0)

    return result