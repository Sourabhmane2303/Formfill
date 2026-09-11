"""
Formfill backend — supports two kinds of "permanent" templates,
auto-detected when you upload one:

  1. content_control - a form that already has Word content controls
     (Developer tab -> Controls), e.g. HR onboarding forms.
  2. jinja - a prose document (letters, legal filings) with {{ field }}
     placeholders typed directly into the text.

Workflow is identical either way:
  POST /templates  - upload the permanent doc ONCE -> schema saved, template_id returned
  POST /generate    - upload only the info doc + template_id -> filled docx
"""
import shutil
import uuid
import json
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document

from sdt_utils import get_schema
import content_control_engine
import legal_engine
import affidavit_engine
from docxtpl import DocxTemplate

APP_DATA = Path(__file__).parent / "data"
TEMPLATES_DIR = APP_DATA / "templates"
JOBS_DIR = APP_DATA / "jobs"
TEMPLATES_INDEX = APP_DATA / "templates_index.json"
FRONTEND_FILE = Path(__file__).parent / "static" / "index.html"

for d in (TEMPLATES_DIR, JOBS_DIR):
    d.mkdir(parents=True, exist_ok=True)
if not TEMPLATES_INDEX.exists():
    TEMPLATES_INDEX.write_text("{}")

app = FastAPI(title="Formfill")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


@app.get("/", response_class=HTMLResponse)
async def frontend():
    if not FRONTEND_FILE.exists():
        raise HTTPException(404, "Frontend not found - expected static/index.html")
    return FRONTEND_FILE.read_text()


def _load_index():
    return json.loads(TEMPLATES_INDEX.read_text())


def _save_index(idx):
    TEMPLATES_INDEX.write_text(json.dumps(idx, indent=2))


def _detect_kind_and_schema(template_path):
    """
    Returns (kind, schema_summary).
    Prefers Jinja detection first (a document could technically contain
    literal '{{' text without intending placeholders, but that's rare;
    content controls are unambiguous when present).
    """
    doc = Document(str(template_path))
    cc_schema = get_schema(doc)
    if cc_schema:
        return "content_control", cc_schema

    tpl = DocxTemplate(str(template_path))
    variables = sorted(tpl.get_undeclared_template_variables())
    if variables:
        return "jinja", [{"variable": v} for v in variables]

    return None, []


@app.post("/templates")
async def upload_template(name: str = Form(...), file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Template must be a .docx file")

    template_id = str(uuid.uuid4())
    template_path = TEMPLATES_DIR / f"{template_id}.docx"
    with open(template_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        kind, schema = _detect_kind_and_schema(template_path)
    except Exception as e:
        template_path.unlink(missing_ok=True)
        raise HTTPException(400, f"Could not parse document: {e}")

    if kind is None:
        template_path.unlink(missing_ok=True)
        raise HTTPException(
            400,
            "No fillable fields found. Either insert Word content controls "
            "(Developer tab > Controls) for each field, or type {{ field_name }} "
            "placeholders directly into the document text.",
        )

    idx = _load_index()
    idx[template_id] = {"name": name, "filename": file.filename, "kind": kind, "schema": schema}
    _save_index(idx)

    return {"template_id": template_id, "name": name, "kind": kind, "fields": schema}


@app.get("/templates")
async def list_templates():
    idx = _load_index()
    return [
        {"template_id": tid, "name": m["name"], "kind": m["kind"], "field_count": len(m["schema"])}
        for tid, m in idx.items()
    ]


@app.get("/templates/{template_id}")
async def get_template(template_id: str):
    idx = _load_index()
    if template_id not in idx:
        raise HTTPException(404, "Template not found")
    return idx[template_id]


@app.delete("/templates/{template_id}")
async def delete_template(template_id: str):
    idx = _load_index()
    if template_id not in idx:
        raise HTTPException(404, "Template not found")
    del idx[template_id]
    _save_index(idx)
    (TEMPLATES_DIR / f"{template_id}.docx").unlink(missing_ok=True)
    return {"status": "deleted"}


@app.post("/generate")
async def generate(template_id: str = Form(...), file: UploadFile = File(...)):
    idx = _load_index()
    if template_id not in idx:
        raise HTTPException(404, "Template not found. Set it up first via /templates.")
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Info doc must be a .docx file")

    kind = idx[template_id]["kind"]
    job_id = str(uuid.uuid4())
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    info_path = job_dir / "info.docx"
    with open(info_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    template_path = TEMPLATES_DIR / f"{template_id}.docx"
    output_path = job_dir / "output.docx"

    engine = content_control_engine if kind == "content_control" else legal_engine
    report = engine.fill_template(str(template_path), str(info_path), str(output_path))
    (job_dir / "report.json").write_text(json.dumps(report, indent=2))

    return {"job_id": job_id, "download_url": f"/generate/{job_id}/download", "report": report}


@app.post("/draft-reply")
async def draft_reply(
    reference_doc: UploadFile = File(..., description="Style exemplar - a finished document with no placeholders"),
    case_info_doc: UploadFile = File(..., description="Structured facts sheet - see case_info_parser.py for expected shape"),
):
    """
    For documents where the variable content is a set of substantive
    paragraphs (not fixed fields) - e.g. affidavits, replies. Unlike
    /templates + /generate, the reference doc here is NOT registered as a
    reusable template: each call both drafts prose (via Groq/Llama) and
    renders the docx in one step, since the "template" (the reference doc's
    style) isn't reused verbatim the way a form's fields are.

    Requires GROQ_API_KEY set in the server environment.
    """
    for f, label in ((reference_doc, "reference_doc"), (case_info_doc, "case_info_doc")):
        if not f.filename.lower().endswith(".docx"):
            raise HTTPException(400, f"{label} must be a .docx file")

    job_id = str(uuid.uuid4())
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    reference_path = job_dir / "reference.docx"
    case_info_path = job_dir / "case_info.docx"
    output_path = job_dir / "output.docx"

    with open(reference_path, "wb") as f:
        shutil.copyfileobj(reference_doc.file, f)
    with open(case_info_path, "wb") as f:
        shutil.copyfileobj(case_info_doc.file, f)

    try:
        _, drafted = affidavit_engine.generate_reply_document(
            str(reference_path), str(case_info_path), str(output_path)
        )
    except ValueError as e:
        # case info doc missing an expected section
        raise HTTPException(400, str(e))
    except RuntimeError as e:
        # missing API key, or model returned something unusable
        raise HTTPException(502, str(e))

    (job_dir / "drafted.json").write_text(json.dumps(drafted, indent=2))

    return {
        "job_id": job_id,
        "download_url": f"/generate/{job_id}/download",
        "drafted": drafted,
        "note": "AI-drafted for a court filing - have it reviewed by a qualified advocate before use.",
    }


@app.get("/generate/{job_id}/download")
async def download(job_id: str):
    output_path = JOBS_DIR / job_id / "output.docx"
    if not output_path.exists():
        raise HTTPException(404, "Job output not found")
    return FileResponse(
        str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="filled_document.docx",
    )


@app.get("/generate/{job_id}/report")
async def get_report(job_id: str):
    report_path = JOBS_DIR / job_id / "report.json"
    if not report_path.exists():
        raise HTTPException(404, "Report not found")
    return json.loads(report_path.read_text())
