# backend/ai_service/ppm_splitter.py
"""
Local, pre-AI page-range locator for Private Placement Memorandums (PPMs).

PPMs typically run 100-300+ pages, but only a handful actually carry the
financial data this platform extracts — the rest is legal/subscription
boilerplate. Sending the whole document to the AI would be slow, expensive,
and — per the historical VPS timeout issues already fixed once for the CRE
pipeline (see engineering-principles: shared hosting's hard nginx timeout
killing long AI calls) — a real risk of tripping the same execution ceiling
all over again on a much bigger document.

This module runs a cheap, local, keyword-anchor scan over the PDF's own text
(via pdfplumber) BEFORE any AI call happens, and narrows the document down to
just the pages likely to hold:
  1. Sources and Uses of Funds
  2. Fee Structure / Sponsor Load Matrix
  3. Historical Operations / Pro Forma

These three sections are NOT assumed to be contiguous — real-world PPMs
scatter them across different parts of the document depending on the
sponsor's template, so each is located independently.

DESIGN PRINCIPLE — fail loud, not silent: this heuristic has not been tuned
against a broad sample of real sponsor templates yet. If it cannot
confidently locate all three sections, it refuses to guess — it raises
PPMSplitterError rather than silently passing through a wrong (or empty)
slice, so a bad extraction doesn't quietly become a wrong number in
someone's underwriting model. The caller (views.py PPMUnderwriteView)
surfaces this as a 400 with needs_manual_page_range=True; the frontend's
fallback card lets the user supply `manual_pages` explicitly instead, which
skips the heuristic entirely.
"""
import pandas as pd
import pdfplumber

from .parsers import FileValidationError

# How many pages after an anchor hit are considered part of that section
# (not counting the anchor page itself). A target of "5 to 8 pages total"
# across 3 sections works out to roughly 1-3 pages per section —
# deliberately generous so a genuine multi-page table isn't truncated, at
# the cost of occasionally pulling in a page of surrounding narrative.
PAGES_AFTER_ANCHOR = 2

# Anchor phrases are matched case-insensitively as substrings against each
# page's extracted text. These are starting points, not a tuned final list —
# expect to extend these once this has run against real sponsor PPMs.
SECTION_ANCHORS = {
    "sources_and_uses": [
        "sources and uses of funds",
        "estimated use of proceeds",
    ],
    "fee_structure": [
        "acquisition fees",
        "sponsor load matrix",
    ],
    "historical_operations": [
        "operating forecast",
        "underwriting assumptions",
    ],
}

SECTION_LABELS = {
    "sources_and_uses": "Sources and Uses of Funds",
    "fee_structure": "Fee Structure / Sponsor Load Matrix",
    "historical_operations": "Historical Operations / Pro Forma",
}


class PPMSplitterError(Exception):
    """
    Raised when the heuristic can't confidently locate all 3 target sections
    and no manual_pages override was given. Carries a user-facing message;
    the caller (views.py) is responsible for surfacing the manual-page-range
    fallback in its response.
    """
    pass


def _find_section_anchor_page(pages_text: list, anchors: list):
    """
    Returns the 0-indexed page number of the FIRST page whose text contains
    any of the given anchor phrases (case-insensitive substring match), or
    None if no page matches.
    """
    for page_idx, text in enumerate(pages_text):
        if not text:
            continue
        lower_text = text.lower()
        if any(anchor in lower_text for anchor in anchors):
            return page_idx
    return None


def locate_ppm_target_pages(ppm_file, manual_pages=None) -> dict:
    """
    Returns {"pages": [0-indexed page numbers, sorted], "sections": {...}}.

    If manual_pages is given (a list of 1-indexed page numbers, as entered
    by the user in the manual-page-range fallback), the heuristic is skipped
    entirely and exactly those pages are used.

    Raises PPMSplitterError if manual_pages is None and one or more of the 3
    target sections can't be located. Raises FileValidationError for
    malformed input (unreadable PDF, out-of-range manual page numbers) —
    same exception type the rest of the parsing pipeline uses for "this
    input is bad", as distinct from "the heuristic couldn't find it".
    """
    ppm_file.seek(0)
    try:
        with pdfplumber.open(ppm_file) as pdf:
            total_pages = len(pdf.pages)

            if manual_pages is not None:
                out_of_range = [p for p in manual_pages if p < 1 or p > total_pages]
                if out_of_range:
                    raise FileValidationError(
                        f"Page number(s) {out_of_range} are out of range for this "
                        f"{total_pages}-page document."
                    )
                return {
                    "pages": sorted({p - 1 for p in manual_pages}),
                    "sections": {"manual_override": True},
                }

            pages_text = [page.extract_text() for page in pdf.pages]
    except FileValidationError:
        raise
    except Exception as e:
        raise FileValidationError(f"Could not read PPM PDF '{ppm_file.name}': {str(e)}")
    finally:
        ppm_file.seek(0)

    located_sections = {}
    missing_sections = []

    for section_key, anchors in SECTION_ANCHORS.items():
        anchor_page = _find_section_anchor_page(pages_text, anchors)
        located_sections[section_key] = anchor_page
        if anchor_page is None:
            missing_sections.append(SECTION_LABELS[section_key])

    if missing_sections:
        raise PPMSplitterError(
            "Could not confidently locate the following required section(s) in "
            f"this document: {', '.join(missing_sections)}. This can happen when "
            "a sponsor uses non-standard section headings. Please manually enter "
            "the page range(s) containing your financial summary tables and "
            "resubmit."
        )

    target_pages = set()
    for anchor_page in located_sections.values():
        for offset in range(PAGES_AFTER_ANCHOR + 1):
            page_idx = anchor_page + offset
            if page_idx < total_pages:
                target_pages.add(page_idx)

    return {"pages": sorted(target_pages), "sections": located_sections}


def extract_ppm_target_text(ppm_file, manual_pages=None):
    """
    Main entry point for the PPM pipeline. Locates the target pages
    (heuristically, or from manual_pages if given), then extracts text +
    tables from ONLY those pages — mirroring extract_text_from_pdf's
    "--- PAGE N ---" marker format from parsers.py, so downstream provenance
    citations (page_location) still line up with real page numbers in the
    source document.

    Returns (extracted_text: str, located_sections_metadata: dict).
    Raises PPMSplitterError or FileValidationError — see
    locate_ppm_target_pages.
    """
    location_result = locate_ppm_target_pages(ppm_file, manual_pages=manual_pages)
    target_page_indices = location_result["pages"]

    ppm_file.seek(0)
    text_content = []
    try:
        with pdfplumber.open(ppm_file) as pdf:
            for page_idx in target_page_indices:
                if page_idx >= len(pdf.pages):
                    continue
                page = pdf.pages[page_idx]

                text = page.extract_text()
                if text:
                    text_content.append(f"--- PAGE {page_idx + 1} ---\n{text}")

                for table in page.extract_tables():
                    if not table or not table[0]:
                        continue
                    df = pd.DataFrame(table[1:], columns=table[0])
                    text_content.append(f"\n[TABLE ON PAGE {page_idx + 1}]\n{df.to_string(index=False)}\n")
    except Exception as e:
        raise FileValidationError(f"Could not extract target pages from '{ppm_file.name}': {str(e)}")
    finally:
        ppm_file.seek(0)

    if not text_content:
        raise FileValidationError(
            f"No extractable text found on the target pages of '{ppm_file.name}'. "
            "These pages may be scanned/image-only and need OCR."
        )

    return "\n".join(text_content), location_result["sections"]
