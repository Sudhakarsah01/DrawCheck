"""
extraction.py
-------------
Everything about pulling text and coordinates out of a drawing PDF.
No Streamlit, no checks, no rendering -- just PyMuPDF plumbing, so it can
be unit-tested and reused on its own.
"""

import fitz  # PyMuPDF


def extract_words(page):
    """List of (rect, word) for every word on the page, in RENDERED page
    space (i.e. already corrected for page rotation). This is the
    coordinate system get_pixmap() renders into, so a rect from here maps
    directly onto the image you show the user."""
    rot = page.rotation_matrix
    return [(fitz.Rect(w[:4]) * rot, w[4]) for w in page.get_text("words")]


def extract_spans(page):
    """List of (rect, text) per text run (a 'span' = one continuous run of
    text as a single PDF text object), also rotation-corrected. Spans keep
    internal spacing intact, unlike word-mode which splits on whitespace --
    that's what makes spacing checks possible."""
    rot = page.rotation_matrix
    raw = page.get_text("dict")
    spans = []
    for block in raw.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if text.strip():
                    spans.append((fitz.Rect(span["bbox"]) * rot, text))
    return spans


def full_page_text(page):
    """The whole sheet's text as one string -- what gets sent to the AI
    check in a single call, instead of looping word by word."""
    return page.get_text("text")


def is_vector_pdf(page):
    """True if the page has real embedded text (CAD export) rather than
    being a scanned raster image with no text layer."""
    return bool(page.get_text().strip())


def clean_word(txt):
    """Return a clean standalone word to spell-check, or None to skip.

    Deliberately conservative: anything containing a digit, a hyphen, a
    slash, or another symbol is treated as a tag/code/ID (e.g.
    'TEF-A-LG-02', '300x200', 'HPSHS-ARA-B00A') and skipped entirely --
    NOT stripped down to its letters and re-glued, which was an earlier
    bug here that turned 'TEF-A-LG-02' into the nonsense string 'TEFALG'
    and fed it into the spelling checks. Skipping non-word tokens outright
    is what lets the checks be aggressive on real words without
    false-flagging every tag on the sheet.
    """
    t = txt.strip(".,:;()[]")
    if not t:
        return None
    if any(c.isdigit() for c in t):
        return None
    if any(ch in t for ch in "-/\\@#$%&*+=_"):
        return None
    if "'" in t:
        return None  # possessives (CHILDREN'S) -- skip, not worth the noise
    return t if t.isalpha() else None