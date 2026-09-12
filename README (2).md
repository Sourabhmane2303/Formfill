# Formfill — draft a reply from a reference document

Upload a finished document with no placeholders as the style guide,
plus a structured facts sheet, and Claude drafts the new document's
prose while the layout is generated deterministically. This is meant
for documents where the content isn't a fixed set of fields — like
affidavits or replies.

> **Only Word `.docx` files are accepted** — both the reference
> document and the facts sheet. PDFs and the legacy `.doc` format are
> not supported; if you have a `.doc` file, open it in Word and save
> it as `.docx` first, and if you have a PDF, recreate it as a
> `.docx`.

## How it works

1. **Upload the reference document** — a finished, real example of
   the kind of document you want (no placeholders), used as the style
   guide for tone, structure, and phrasing.
2. **Upload the case information / facts sheet** — a structured
   document describing the specific facts of the new case.
3. Claude drafts the new document's opening, body paragraphs, and any
   prayer/relief items based on the reference's style and the
   supplied facts. The result is assembled into a `.docx` and
   returned, along with a breakdown of what was drafted for each
   section.

## Project layout

```
backend/
  main.py                    FastAPI app
  draft_engine.py             Drafts new prose (via Claude/Groq) from a reference doc + facts sheet
  affidavit_engine.py         Fill/draft engine specific to affidavit-style documents
  case_info_parser.py         Parses the case information / facts sheet for drafting
  docx_builder.py             Assembles the final .docx output from the drafted content
  requirements.txt
  .env                        Holds GROQ_API_KEY (not committed — create this yourself)
frontend/
  index.html                  Single-page UI (no build step) — draft-a-reply only
test/
  affidavit_template.docx     Sample reference document
  case_info.docx              Sample facts sheet
```

## Running it

1. **Get the code.** Either clone the repo:
   ```bash
   git clone <repo-url>
   ```
   or download the ZIP from GitHub ("Code" → "Download ZIP") and
   extract it.

2. **Open the project in an IDE.** VS Code or any other editor works —
   open the extracted/cloned `Formfill-main` folder.

3. **Open a terminal in the IDE** and move into the project folder:
   ```bash
   cd Formfill-main
   ```

4. **Get a Groq API key.** Go to
   [console.groq.com/keys](https://console.groq.com/keys), sign in,
   and create a new API key.

5. **Add the key to a `.env` file.** In the `backend/` folder, create
   a file named `.env` (if it doesn't already exist) and add:
   ```
   GROQ_API_KEY=""
   ```
   Paste your key between the quotes.

6. **Install dependencies and start the backend:**
   ```bash
   cd backend
   pip install -r requirements.txt
   uvicorn main:app --reload --port 8000
   ```

7. **Open the frontend.** Open `frontend/index.html` directly in any
   browser (it talks to `http://localhost:8000`).

8. **Use the app.** Select your two files — the reference document
   and the case information / facts sheet — and get the drafted
   document back. Both files must be Word `.docx` files (PDFs and
   legacy `.doc` files aren't accepted).

## API

| Method | Path | Purpose |
|---|---|---|
| POST | `/draft-reply` | Upload the reference doc (`reference_doc`) and facts sheet (`case_info_doc`). Returns the drafted sections and a `download_url` for the assembled `.docx`. |

Generated documents live under `backend/data/` (created
automatically) — swap for a real database/blob storage before
production use.

## Things to adapt for your real documents

- **Facts sheet format**: `case_info_parser.py` is the place to extend
  if your real facts sheets are structured differently than the
  sample.
- **Drafting behavior**: `draft_engine.py` controls the prompt/logic
  used to generate the opening, body paragraphs, and prayer items —
  adjust it to match the tone and structure of your own document
  types.
- **Output formatting**: `docx_builder.py` controls how the drafted
  sections are assembled into the final `.docx` layout.

**Note:** `main.py`, `draft_engine.py`, `affidavit_engine.py`, and
`case_info_parser.py` reflect the current draft-only backend. If you
still have older files from the save-form/fill-form workflow
(`content_control_engine.py`, `legal_engine.py`, `info_parser.py`,
`field_mapper.py`, `sdt_utils.py`) lying around in your project, they
are no longer used and can be removed.
