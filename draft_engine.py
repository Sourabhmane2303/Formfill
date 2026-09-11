"""
Drafts the variable-length prose portions of a legal reply document
(opening deposition line, numbered paragraphs, prayer sentences) by
prompting an LLM with:
  1. the full text of a reference document as a style/phrasing exemplar
  2. the structured facts for the new case (from case_info_parser)

Everything that ISN'T prose (party names, case number, dates, headings,
table layout) is handled deterministically elsewhere (docx_builder.py) -
the model is only asked to draft what genuinely requires drafting: turning
bullet-point facts into properly worded legal paragraphs in a matching
register.

Requires GROQ_API_KEY in the environment. Model is configurable via
GROQ_MODEL (defaults to Llama 3.3 70B, which handles formal legal register
noticeably better than the smaller Groq models).

API key setup:
    Put GROQ_API_KEY=your-key-here in a .env file next to main.py (add
    .env to .gitignore), or export it in your shell before starting the
    server. This module auto-loads a .env file on import if python-dotenv
    is installed; if not, it just falls back to whatever is already in
    the environment.
"""
import os
import json
import re
import time
import logging
import requests
from docx import Document

try:
    from dotenv import load_dotenv
    load_dotenv()  # populates os.environ from a .env file, if one exists
except ImportError:
    # python-dotenv not installed - fine, we just rely on the environment
    # already having GROQ_API_KEY set (e.g. exported in the shell, or
    # injected by whatever process manager/deployment platform is in use).
    pass

logger = logging.getLogger(__name__)

API_URL = "https://api.groq.com/openai/v1/chat/completions"
DEFAULT_MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-120b")
MAX_JSON_RETRIES = 2  # extra attempts if the model returns malformed JSON


class GroqError(RuntimeError):
    pass


class GroqAuthError(GroqError):
    pass


class GroqRateLimitError(GroqError):
    pass


class GroqDraftError(GroqError):
    pass


def _extract_reference_text(reference_docx_path):
    doc = Document(reference_docx_path)
    parts = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _build_prompt(reference_text, context):
    deponent = context['deponent']
    case = context['case']
    points = context['reply_points']

    points_block = "\n".join(
        f"{p['title']}\n" + "\n".join(f"  - {b}" for b in p['bullets'])
        for p in points
    )
    prayer_block = "\n".join(f"- {item}" for item in context['prayer_items'])

    return f"""You are drafting the prose portions of a legal Affidavit in Reply for an \
Indian court filing, in the exact structure and phrasing register of the reference \
affidavit below. Use it as your style guide for tone, sentence construction, and \
standard legal phrases (e.g. "I say that...", "save and except those specifically \
admitted herein", "misconceived, devoid of merits", "liable to be dismissed in \
limine", "hereto annexed and marked as EXHIBIT-'X'").

=== REFERENCE AFFIDAVIT (style exemplar only - do not reuse its facts) ===
{reference_text}
=== END REFERENCE ===

Now draft the new affidavit's prose using ONLY the facts below. Do not invent any \
fact, name, date, or relief not given here. Do not add extra reliefs to the prayer \
beyond what is listed, though standard catch-all closing language (e.g. "and such \
other reliefs as the Court may deem fit") is acceptable if it fits the style.

Important constraints:
- Do not invent any fact, name, date, exhibit number, or relief not explicitly given above.
- Do not add extra prayer items beyond those listed; you may only add standard catch-all
  language (e.g. "and such other reliefs as this Hon'ble Court may deem fit and proper")
  if it matches the reference style.
- The "paragraphs" array must have exactly {len(points)} entries, one per reply point.

DEPONENT: {json.dumps(deponent, ensure_ascii=False)}
CASE: {json.dumps(case, ensure_ascii=False)}

REPLY POINTS (each must become exactly one numbered paragraph, in this order):
{points_block}

PRAYER ITEMS:
{prayer_block}

Respond with ONLY valid JSON (no markdown fences, no commentary), in this exact shape:
{{
  "opening": "<the single opening deposition sentence, e.g. 'I, <name>, <designation/role>, ... do hereby <verb> and state as under:'>",
  "paragraphs": ["<paragraph 1 text, no leading number>", "<paragraph 2 text>", ...],
  "prayer_intro": "<the lead-in sentence before the lettered prayer list>",
  "prayer_items": ["<prayer item a text, no leading letter>", ...]
}}

The "paragraphs" array must have exactly {len(points)} entries, one per reply point, \
in the same order."""


def _call_groq(prompt, api_key, model):
    session = getattr(_call_groq, "_session", None)
    if session is None:
        session = requests.Session()
        _call_groq._session = session

    def _do_request():
        return session.post(
            API_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": model or DEFAULT_MODEL,
                "max_tokens": 2000,
                "temperature": 0.2,
                "response_format": {"type": "json_object"},
                "messages": [
                    {
                        "role": "system",
                        "content": "You output ONLY raw JSON. No markdown fences, no commentary, "
                                   "no text before or after the JSON object.",
                    },
                    {"role": "user", "content": prompt},
                ],
            },
            timeout=60,
        )

    # Network-level failures (unreachable host, DNS failure, connection
    # refused, timeout, etc.) raise requests.exceptions.RequestException,
    # NOT an HTTP status. Those are not ValueError/RuntimeError, so if left
    # uncaught here they fall straight through main.py's except clauses and
    # surface to the client as a bare, unhandled 500. Wrap every _do_request()
    # call so every failure mode becomes a GroqError (a RuntimeError subclass)
    # that main.py's `except RuntimeError` can turn into a clean 502.
    try:
        resp = _do_request()
    except requests.exceptions.RequestException as e:
        raise GroqError(f"Could not reach Groq API: {e}") from e

    if resp.status_code == 401:
        raise GroqAuthError("Groq API authentication failed. Check your API key.")

    if resp.status_code == 429:
        retry_after = float(resp.headers.get("Retry-After", 5))
        time.sleep(retry_after)
        try:
            resp = _do_request()
        except requests.exceptions.RequestException as e:
            raise GroqError(f"Could not reach Groq API on retry: {e}") from e
        if resp.status_code == 429:
            raise GroqRateLimitError(
                f"Groq rate limit exceeded. Retry-After: {retry_after}s"
            )

    if resp.status_code >= 500:
        time.sleep(2)
        try:
            resp = _do_request()
        except requests.exceptions.RequestException as e:
            raise GroqError(f"Could not reach Groq API on retry: {e}") from e

    try:
        resp.raise_for_status()
    except requests.exceptions.HTTPError as e:
        raise GroqError(f"Groq API returned an error: {e}") from e

    try:
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError) as e:
        # ValueError here is resp.json()'s JSONDecodeError subclass, not the
        # case-info ValueError main.py checks for - but since main.py's
        # /draft-reply only distinguishes ValueError (case-info problem) from
        # RuntimeError (API/model problem), letting a JSON-decode ValueError
        # escape here would be *mis-caught* as a 400 case-info error instead
        # of the 502 API error it actually is. Re-wrap as GroqError so it
        # lands in the right bucket.
        raise GroqError(f"Unexpected response shape from Groq API: {e}") from e


def _validate_draft(drafted, expected_paragraphs, expected_prayer_items):
    required_keys = {"opening", "paragraphs", "prayer_intro", "prayer_items"}
    if set(drafted.keys()) != required_keys:
        missing = required_keys - set(drafted.keys())
        extra = set(drafted.keys()) - required_keys
        raise ValueError(f"Invalid draft schema. Missing: {missing}, Extra: {extra}")

    if not isinstance(drafted["opening"], str) or not drafted["opening"].strip():
        raise ValueError("'opening' must be a non-empty string.")

    if not isinstance(drafted["paragraphs"], list):
        raise ValueError("'paragraphs' must be a list.")

    if len(drafted["paragraphs"]) != expected_paragraphs:
        raise ValueError(
            f"Expected {expected_paragraphs} paragraphs, got {len(drafted['paragraphs'])}."
        )

    if not isinstance(drafted["prayer_intro"], str):
        raise ValueError("'prayer_intro' must be a string.")

    if not isinstance(drafted["prayer_items"], list):
        raise ValueError("'prayer_items' must be a list.")

    if len(drafted["prayer_items"]) != expected_prayer_items:
        raise ValueError(
            f"Expected {expected_prayer_items} prayer items, got {len(drafted['prayer_items'])}."
        )


def draft_affidavit_prose(reference_docx_path, context, api_key=None, model=None):
    api_key = api_key or os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise GroqAuthError(
            "No API key found. Set GROQ_API_KEY in a .env file or in the "
            "environment, or pass api_key= explicitly."
        )

    reference_text = _extract_reference_text(reference_docx_path)
    prompt = _build_prompt(reference_text, context)
    expected_paragraphs = len(context['reply_points'])
    expected_prayer_items = len(context['prayer_items'])

    last_error = None
    for attempt in range(1 + MAX_JSON_RETRIES):
        logger.info("Groq call attempt %d/%d", attempt + 1, 1 + MAX_JSON_RETRIES)
        text = _call_groq(prompt, api_key, model)

        cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
        try:
            drafted = json.loads(cleaned)
        except json.JSONDecodeError as e:
            last_error = f"Model did not return valid JSON: {e}\nRaw output:\n{text}"
            logger.warning("Attempt %d failed: invalid JSON", attempt + 1)
            continue

        try:
            _validate_draft(drafted, expected_paragraphs, expected_prayer_items)
        except ValueError as e:
            last_error = str(e)
            logger.warning("Attempt %d failed: schema validation error - %s", attempt + 1, last_error)
            continue

        return drafted

    raise GroqDraftError(
        f"Failed after {1 + MAX_JSON_RETRIES} attempt(s). Last error: {last_error}"
    )


if __name__ == '__main__':
    import sys
    from case_info_parser import parse_case_info, to_affidavit_context

    reference_path, case_info_path = sys.argv[1], sys.argv[2]
    ctx = to_affidavit_context(parse_case_info(case_info_path))
    drafted = draft_affidavit_prose(reference_path, ctx)
    print(json.dumps(drafted, indent=2, ensure_ascii=False))