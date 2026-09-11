"""
Third engine for the Formfill backend, alongside content_control_engine.py
(form fields) and legal_engine.py (fixed jinja placeholders).

This one is for documents where the variable content isn't a fixed set of
fields but a variable-length set of substantive paragraphs that must be
*drafted*, not just substituted - e.g. affidavits, replies, and similar
legal filings built from a reference document's style plus a structured
facts sheet.

Pipeline:
  reference.docx (style exemplar, NOT a fillable template)
       |
  case_info.docx (structured facts + variable-length point list)
       |
       v
  case_info_parser.parse_case_info()      -> generic section tree
  case_info_parser.to_affidavit_context() -> flattened fields
       |
       v
  draft_engine.draft_affidavit_prose()    -> LLM drafts ONLY the prose
       |
       v
  docx_builder.build_affidavit()          -> deterministic docx render
"""
from case_info_parser import parse_case_info, to_affidavit_context
from draft_engine import draft_affidavit_prose
from docx_builder import build_affidavit


def generate_reply_document(reference_path, case_info_path, output_path,
                             api_key=None, model=None):
    """Runs the full pipeline. Returns (output_path, drafted_prose_dict) so
    callers can also surface what was drafted for review before filing."""
    parsed = parse_case_info(case_info_path)
    context = to_affidavit_context(parsed)

    drafted = draft_affidavit_prose(reference_path, context, api_key=api_key, model=model)

    build_affidavit(context, drafted, output_path)
    return output_path, drafted


if __name__ == '__main__':
    import sys
    reference_path, case_info_path, output_path = sys.argv[1], sys.argv[2], sys.argv[3]
    out, drafted = generate_reply_document(reference_path, case_info_path, output_path)
    print(f"Wrote {out}")
