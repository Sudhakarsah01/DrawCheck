# Drawing Text QA

Flags spelling and spacing issues on drawing PDFs (tested on CAD-exported
vector PDFs — a scanned/raster drawing needs OCR, not wired up yet).

## Setup

```powershell
pip install -r requirements.txt
```

**AI check (optional, on by default) needs a Groq API key.** Don't paste it
into `app.py` — set it once outside the code:

```powershell
mkdir .streamlit -Force
python -c "open('.streamlit/secrets.toml','w',encoding='utf-8').write('GROQ_API_KEY = \"gsk_your_real_key\"\n')"
```

(Use `python -c` rather than PowerShell's `echo >`, which writes UTF-16
and breaks the TOML parser.)

Run it:

```powershell
streamlit run app.py
```

## Project layout

| File | What it does |
|---|---|
| `extraction.py` | Pulls text + coordinates out of the PDF (PyMuPDF). Also the word tokenizer — deliberately conservative, skips anything with a digit/hyphen/slash so tags like `TEF-A-LG-02` never get mangled into a fake word. |
| `checks.py` | Three local, no-API checks: glossary fuzzy-match, general dictionary, and consistency-based spacing (only flags spacing that *differs* from how the same text is spaced elsewhere on the sheet — not every double space). |
| `ai_check.py` | Optional fourth check: one Groq API call per page reviewing the whole sheet's text at once. Key loading/masking/validation lives here too. |
| `render.py` | Draws the flags onto the page image. |
| `app.py` | Streamlit UI only — wires the modules above together. No detection logic lives here. |

## Why three-plus checks instead of one

Each catches something the others miss:
- **Glossary** catches typos *of* terms you've told it about (`RRCHITECT` → `ARCHITECT`)
- **Dictionary** catches ordinary misspellings anywhere, even words nowhere near a glossary term
- **AI check** catches context-aware issues neither rule-based layer can (e.g. a real word used wrong)
- Neither glossary nor dictionary alone is complete: a real-word-but-wrong typo like `PURPOSED` (a valid English word, just wrong here) is invisible to the dictionary layer and only caught by the glossary layer knowing `PROPOSED` is expected.

## Troubleshooting

**401 Invalid API Key** — the app's sidebar shows exactly which key it's
using and where it came from (`secrets.toml` vs environment variable vs
manual entry), plus whether it *looks* like a real key (`gsk_...`, 50+
chars). If it says "environment variable" and the key looks wrong, you
likely have a leftover/placeholder value set in that terminal session
from an earlier test — close the terminal, open a new one, and re-run.

**404 model_not_found** — Groq retires/renames models over time. The
model field in the sidebar is editable at runtime; check
https://console.groq.com/docs/models for the current list if the default
(`meta-llama/llama-4-scout-17b-16e-instruct`) stops working.

**"No module named 'spellchecker'"** — `pip install -r requirements.txt`
(the PyPI package is `pyspellchecker`, but the import name is
`spellchecker` — no "py" prefix).