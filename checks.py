"""
checks.py
---------
The three rule-based/local checks (no API key needed), each with a
deliberately different job so they catch different failure modes:

  glossary_check    -- typos OF known trade terms you've told it about
                        (RRCHITECT vs ARCHITECT)
  dictionary_check   -- ordinary misspellings anywhere on the sheet, even
                        words nowhere near a glossary term
  spacing_check      -- flags spacing only where it's INCONSISTENT with
                        how the same text is spaced elsewhere on the same
                        sheet -- not just because a double space exists.
                        A note that's double-spaced everywhere is a style
                        choice, not an error.
"""

import re
import difflib
from collections import Counter

from spellchecker import SpellChecker
from extraction import clean_word

_SPELLCHECKER = SpellChecker()


def glossary_check(words, glossary, ignored, cutoff=0.8, min_len=4):
    flags = []
    glossary_upper = {g.upper() for g in glossary}
    for rect, txt in words:
        w = clean_word(txt)
        if not w or len(w) < min_len:
            continue
        wu = w.upper()
        if wu in glossary_upper or wu in ignored:
            continue
        match = difflib.get_close_matches(wu, glossary_upper, n=1, cutoff=cutoff)
        if match:
            flags.append({"rect": rect, "word": txt, "suggestion": match[0], "source": "rules"})
    return flags


def dictionary_check(words, glossary, ignored, min_len=5):
    """Complementary to glossary_check: a real-word-but-wrong-in-context
    typo like PURPOSED (a valid English word) won't trigger here, since
    it's not unknown to the dictionary -- that's precisely why both
    checks need to run together. min_len=5 keeps this off short
    acronyms/initials, where a general English dictionary is least
    reliable on a trade drawing."""
    glossary_upper = {g.upper() for g in glossary}
    flags = []
    for rect, txt in words:
        w = clean_word(txt)
        if not w or len(w) < min_len:
            continue
        wu = w.upper()
        if wu in glossary_upper or wu in ignored:
            continue
        if _SPELLCHECKER.unknown([w.lower()]):
            cands = _SPELLCHECKER.candidates(w.lower())
            if cands:
                best = sorted(cands, key=lambda c: abs(len(c) - len(w)))[0]
                flags.append({"rect": rect, "word": txt, "suggestion": best.upper(),
                              "source": "dictionary"})
    return flags


def diff_highlight(wrong, correct):
    """Wrap the differing letters between two words in **markdown bold**,
    so a reviewer sees exactly what's off instead of eyeballing two long
    strings for a difference."""
    sm = difflib.SequenceMatcher(a=wrong.upper(), b=correct.upper())
    w_out, c_out = "", ""
    for tag, a0, a1, b0, b1 in sm.get_opcodes():
        seg_w, seg_c = wrong[a0:a1], correct[b0:b1]
        if tag == "equal":
            w_out += seg_w
            c_out += seg_c
        else:
            if seg_w:
                w_out += f"**{seg_w}**"
            if seg_c:
                c_out += f"**{seg_c}**"
    return w_out, c_out


def _nospace_key(s):
    """All whitespace removed, uppercased -- groups spans that are the
    'same' underlying text regardless of how many spaces separate the
    words, so different spacing variants of one note can be compared."""
    return re.sub(r"\s+", "", s).upper()


def spacing_check(spans, ignored_texts=None, label_max_len=60):
    """See module docstring. Two tiers:
      1. (main signal) same text, more than one spacing variant on the
         sheet -> flag the minority variant(s), name the majority one.
      2. (weak signal) 3+ consecutive spaces in a one-off span with no
         sibling to compare against -- unusual on its own, kept clearly
         lower-confidence than tier 1.
    """
    ignored_texts = ignored_texts or set()
    candidates = []
    for rect, text in spans:
        stripped = text.strip()
        if not stripped or len(stripped) > label_max_len:
            continue
        if stripped.upper() in ignored_texts:
            continue
        candidates.append((rect, stripped, text.rstrip()))

    groups = {}
    for rect, stripped, interior in candidates:
        groups.setdefault(_nospace_key(stripped), []).append((rect, stripped, interior))

    issues = []
    flagged_rect_ids = set()

    for occurrences in groups.values():
        variants = Counter(s for _, s, _ in occurrences)
        if len(variants) > 1:
            majority_form, _ = variants.most_common(1)[0]
            for rect, stripped, interior in occurrences:
                if stripped == majority_form:
                    continue
                issues.append({
                    "rect": rect, "type": "spacing_inconsistency", "text": stripped,
                    "marked": stripped,
                    "detail": f"Spacing differs from the sheet's more common form "
                              f"of this note: '{majority_form}'",
                    "source": "rules",
                })
                flagged_rect_ids.add(id(rect))

    for rect, stripped, interior in candidates:
        if id(rect) in flagged_rect_ids:
            continue
        m = re.search(r"\s{3,}", interior)
        if m:
            i, j = m.span()
            marked = interior[:i] + "[···]" + interior[j:]
            issues.append({"rect": rect, "type": "wide_space", "text": stripped,
                            "marked": marked, "detail": "Unusually wide gap (3+ spaces) — review",
                            "source": "rules"})

    return issues