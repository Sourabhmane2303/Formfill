"""
Engine for templates that use Word content controls (form fields) —
e.g. onboarding forms, application forms. See legal_engine.py for the
Jinja-placeholder engine used for prose/legal documents.
"""
from docx import Document
from sdt_utils import iter_content_controls, get_schema
from info_parser import extract_fields
from field_mapper import map_fields


def extract_template_schema(template_path):
    doc = Document(template_path)
    return get_schema(doc)


def fill_template(template_path, info_doc_path, output_path, score_threshold=60):
    schema = extract_template_schema(template_path)
    extracted = extract_fields(info_doc_path)

    # candidate id = control index; aliases = its tag and title
    candidates = {i: [f.get('tag'), f.get('title')] for i, f in enumerate(schema)}
    mapping, matched_report, unmatched, unfilled_ids = map_fields(
        extracted, candidates, score_threshold=score_threshold
    )

    doc = Document(template_path)
    for idx, cc in enumerate(iter_content_controls(doc)):
        if idx in mapping:
            cc.set_text(mapping[idx])

    doc.save(output_path)

    unfilled_controls = [
        {'tag': schema[i].get('tag'), 'title': schema[i].get('title')}
        for i in unfilled_ids
    ]
    # rename 'matched_field' (a control index) to matched_tag/matched_title for API clarity
    for m in matched_report:
        idx = m.pop('matched_field')
        m['matched_tag'] = schema[idx].get('tag')
        m['matched_title'] = schema[idx].get('title')

    return {
        'filled': matched_report,
        'unmatched_info_fields': unmatched,
        'unfilled_fields': unfilled_controls,
    }
