"""
ai_check.py
-----------
Optional third check: one AI (Groq) call reviewing the WHOLE page's text
at once, for context-aware catches the local checks miss. Deliberately
ONE request per page, not one per word -- looping would multiply latency
and token cost for no accuracy benefit.

Model: openai/gpt-oss-20b -- set from this key's actual confirmed access
(via the "List models available to this key" button in the sidebar,
which calls Groq's /v1/models endpoint). Different Groq accounts/keys
can have access to different model sets, so don't assume any model name
found in docs or elsewhere works universally -- if this ever 404s again,
re-run that button rather than guessing a new name from search results.

--- API key ---
Never hardcode a real key into source you might share or commit. Set it
once outside the code:
  a) .streamlit/secrets.toml next to app.py:   GROQ_API_KEY = "gsk_..."
  b) or an environment variable:               GROQ_API_KEY=gsk_...
get_stored_key() reads either automatically and reports which one it
used, which is the fastest way to diagnose a stale/placeholder key.
"""

import os
import json

DEFAULT_MODEL = "openai/gpt-oss-20b"

SYSTEM_PROMPT = """You are reviewing text extracted from an engineering \
(HVAC/mechanical) drawing sheet for typos and spacing mistakes only.

Rules:
- Equipment tags, codes and dimensions (e.g. "OAF-A-LG-01", "300x200", "200Ø") are \
NEVER errors. Do not flag them.
- Only flag a word if it looks like a genuine misspelling of an ordinary English or \
trade word (e.g. "RRCHITECT" for "ARCHITECT", "PURPOSED" for "PROPOSED").
- Only flag spacing if the SAME phrase appears elsewhere on the sheet with different \
spacing (an inconsistency), or there's an unusually wide gap (3+ spaces). A double \
space used consistently everywhere is a style choice, not an error -- do not flag it.
- Known-correct trade terms for this project: {glossary}
- Return ONLY a JSON object: {{"issues": [{{"found": "...", "suggestion": "...", \
"type": "spelling"|"spacing", "reason": "short reason"}}]}}. No prose, no markdown \
fences. If there are no issues, return {{"issues": []}}.
"""


def get_stored_key(streamlit_secrets=None):
    """Read the key from secrets.toml or an env var, and report WHICH one
    it came from -- (key_or_None, source_string). A 401 that survives
    fixing secrets.toml is almost always: (1) secrets.toml not found from
    the working directory streamlit was launched in, so it silently falls
    through to (2) a leftover/placeholder environment variable from an
    earlier test command in the same terminal session. Showing the source
    makes that immediately diagnosable instead of guessed at."""
    key, source = None, None
    if streamlit_secrets is not None:
        try:
            if "GROQ_API_KEY" in streamlit_secrets:
                key, source = streamlit_secrets["GROQ_API_KEY"], "secrets.toml"
        except Exception:
            pass
    if not key:
        env_key = os.environ.get("GROQ_API_KEY")
        if env_key:
            key, source = env_key, "environment variable"
    if key:
        key = key.strip().strip('"').strip("'")
    return (key or None), source


def mask_key(key):
    """Enough of the key to confirm it's the right one without showing
    the whole secret -- e.g. 'gsk_ab...wxyz (56 chars)'."""
    if not key:
        return "(none)"
    if len(key) <= 10:
        return f"{key[:2]}...({len(key)} chars)"
    return f"{key[:6]}...{key[-4:]} ({len(key)} chars)"


def key_looks_valid(key):
    return bool(key) and key.startswith("gsk_") and len(key) >= 40


def list_available_models(api_key):
    """Ask Groq directly which models THIS key can see. The most reliable
    way to resolve repeated 404s across multiple model names -- if
    several different models (including ones known to be current
    production models) all 404 for the same key, it's very unlikely to be
    a model-name problem and much more likely something account-specific
    (a restricted-scope key, a free-tier limitation, etc). This settles
    it either way instead of guessing more names."""
    from groq import Groq

    client = Groq(api_key=api_key)
    resp = client.models.list()
    return [m.id for m in resp.data]


def run_check(full_text, glossary, api_key, model=DEFAULT_MODEL):
    """Single chat-completion call reviewing the WHOLE page's text at
    once. Returns a list of issue dicts. Raises on API/parse failure so
    the caller can show a clear error instead of silently returning
    nothing.

    reasoning_format="hidden" and a low reasoning_effort matter here:
    openai/gpt-oss-20b (and other Groq 'reasoning' models) spend part of
    their token budget on internal chain-of-thought before the actual
    answer. Left uncontrolled, that reasoning can consume the whole
    response and the JSON payload never gets emitted -- which is exactly
    what an empty `failed_generation` in a json_validate_failed error
    means. Hiding the reasoning and capping its effort keeps the budget
    on the actual JSON output. max_completion_tokens is set generously
    since a full sheet's text plus a JSON issue list can be sizeable.
    These reasoning params are ignored by non-reasoning models, so this
    stays safe if the model is switched later.
    """
    from groq import Groq

    client = Groq(api_key=api_key)
    system = SYSTEM_PROMPT.format(glossary=", ".join(sorted(glossary)))
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": full_text},
        ],
        response_format={"type": "json_object"},
        temperature=0,
        max_completion_tokens=4096,
        reasoning_effort="low",
        reasoning_format="hidden",
    )
    content = resp.choices[0].message.content
    if not content or not content.strip():
        raise ValueError(
            "AI returned an empty response (no text to parse as JSON). "
            "This usually means the model ran out of its token budget "
            "before producing output -- try again, or pick a smaller/"
            "faster model from the available-models list."
        )
    return json.loads(content).get("issues", [])


def locate_issue(found_text, words, spans):
    """Best-effort: find a bbox for an AI-reported string by matching it
    against extracted words/spans. Returns None if not found -- caller
    still lists the issue, just without a marker on the drawing."""
    ft = found_text.strip().upper()
    for rect, txt in words:
        if txt.strip().upper() == ft:
            return rect
    for rect, txt in spans:
        if ft in txt.strip().upper():
            return rect
    return                                                                         