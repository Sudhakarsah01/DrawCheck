"""
analyze.py
----------
Runs every check (glossary, dictionary, AI) across every page of the
document in one pass, producing a single report. Deliberately separate
from app.py's rendering so it can run behind an explicit "Analyze" button
instead of firing automatically on upload or on every widget interaction
-- important because the AI check costs real API time/tokens per page,
and nobody wants that firing just because they nudged a page-number field.
"""

from extraction import extract_words, extract_spans, full_page_text, is_vector_pdf
from checks import glossary_check, dictionary_check, spacing_check
from ai_check import run_check, locate_issue


def analyze_document(doc, *, glossary, ignored_words, ignored_spacing,
                      use_glossary, use_dictionary, use_ai, ai_key, ai_model,
                      cutoff=0.8, progress_cb=None):
    """Analyze every page. Returns:
        {
          "pages": {page_num: {"spell_flags": [...], "spacing_flags": [...],
                                "scanned": bool}},
          "ai_errors": [(page_num, error_message), ...],
        }
    progress_cb(current_page_index, total_pages, page_label) is called
    before each page starts, so the caller can drive a progress bar.
    """
    results = {"pages": {}, "ai_errors": []}
    total = len(doc)

    for pno in range(total):
        if progress_cb:
            progress_cb(pno, total, f"page {pno + 1}/{total}")

        page = doc[pno]

        if not is_vector_pdf(page):
            results["pages"][pno] = {"spell_flags": [], "spacing_flags": [], "scanned": True}
            continue

        words = extract_words(page)
        spans = extract_spans(page)

        spell_flags, spacing_flags = [], []
        if use_glossary:
            spell_flags += glossary_check(words, glossary, ignored_words, cutoff)
        if use_dictionary:
            spell_flags += dictionary_check(words, glossary, ignored_words)
        spacing_flags += spacing_check(spans, ignored_spacing)

        if use_ai and ai_key:
            try:
                ai_issues = run_check(full_page_text(page), glossary, ai_key, ai_model)
                for it in ai_issues:
                    rect = locate_issue(it.get("found", ""), words, spans)
                    if it.get("type") == "spacing":
                        spacing_flags.append({
                            "rect": rect, "type": "ai_spacing", "text": it.get("found", ""),
                            "marked": it.get("found", ""),
                            "detail": it.get("reason", "Spacing issue (AI)"), "source": "llm",
                        })
                    else:
                        spell_flags.append({
                            "rect": rect, "word": it.get("found", ""),
                            "suggestion": it.get("suggestion", ""), "source": "llm",
                        })
            except Exception as e:
                results["ai_errors"].append((pno, str(e)))

        results["pages"][pno] = {
            "spell_flags": spell_flags, "spacing_flags": spacing_flags, "scanned": False,
        }

    return results


def settings_fingerprint(*, file_hash, glossary, use_glossary, use_dictionary,
                          use_ai, ai_model, cutoff):
    """A hashable snapshot of everything that affects the result, so the UI
    can tell whether a cached report is stale (settings changed since the
    last Analyze click) without re-running anything."""
    return (file_hash, tuple(sorted(glossary)), use_glossary, use_dictionary,
            use_ai, ai_model, cutoff)