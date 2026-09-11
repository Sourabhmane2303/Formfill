"""
Engine for templates that use {{ jinja_placeholders }} embedded directly
in the document prose — used for legal documents, letters, and other
free-flowing text where content controls don't fit naturally.
"""
from xml.sax.saxutils import escape
from docxtpl import DocxTemplate
from info_parser import extract_fields
from field_mapper import map_fields


def extract_template_variables(template_path):
    tpl = DocxTemplate(template_path)
    return sorted(tpl.get_undeclared_template_variables())


def fill_template(template_path, info_doc_path, output_path, score_threshold=60):
    variable_names = extract_template_variables(template_path)
    extracted = extract_fields(info_doc_path)

    candidates = {name: [name] for name in variable_names}
    mapping, matched_report, unmatched, unfilled_ids = map_fields(
        extracted, candidates, score_threshold=score_threshold
    )

    # docxtpl builds the merged XML as a raw string before parsing it,
    # so unescaped '&', '<', '>' in a value would corrupt (or silently
    # drop from) the output. Escape every value before rendering, and
    # fill any unmatched variable with '' so rendering doesn't error.
    render_context = {name: escape(mapping.get(name, '')) for name in variable_names}

    tpl = DocxTemplate(template_path)
    tpl.render(render_context)
    tpl.save(output_path)

    for m in matched_report:
        m['matched_variable'] = m.pop('matched_field')

    return {
        'filled': matched_report,
        'unmatched_info_fields': unmatched,
        'unfilled_fields': [{'variable': v} for v in unfilled_ids],
    }
