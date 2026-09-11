"""
Parses a structured 'Case Information' docx into a plain dict.

Expected shape (seen consistently in practice):
  - Heading 1: document title
  - Heading 2 sections, each EITHER:
      a) followed immediately by a 2-column table -> parsed as {key: value}
      b) followed by a flat list of bullet paragraphs (style 'List Paragraph')
         -> parsed as a list of strings (e.g. the Prayer section)
      c) followed by Heading 3 sub-points, each with its own bullet list
         -> parsed as [{'title': ..., 'bullets': [...]}, ...] (e.g. Reply Points)

This walks the document body in true reading order (paragraphs interleaved
with tables), so section->table/list association is positional, not
assumed from a fixed template.
"""
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn


def _iter_block_items(doc):
    """Yield each top-level paragraph/table in the document body, in order."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)


def _table_to_dict(table):
    out = {}
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells]
        if len(cells) >= 2 and cells[0]:
            out[cells[0]] = cells[1]
    return out


def _slugify(heading_text):
    return heading_text.strip().lower().replace(' ', '_').replace('-', '_')


def parse_case_info(docx_path):
    doc = Document(docx_path)
    items = list(_iter_block_items(doc))

    sections = {}  # slug -> parsed content
    section_order = []
    i = 0
    n = len(items)
    current_h2 = None

    while i < n:
        item = items[i]

        if isinstance(item, Paragraph):
            style = item.style.name if item.style else ''
            text = item.text.strip()

            if style == 'Heading 2' and text:
                current_h2 = text
                section_order.append(current_h2)
                i += 1

                # Look ahead: table, flat bullet list, or H3 sub-points
                collected_bullets = []
                subpoints = []
                current_subpoint = None

                while i < n:
                    nxt = items[i]

                    if isinstance(nxt, Table):
                        sections[current_h2] = {'type': 'table', 'data': _table_to_dict(nxt)}
                        i += 1
                        break

                    if isinstance(nxt, Paragraph):
                        nstyle = nxt.style.name if nxt.style else ''
                        ntext = nxt.text.strip()

                        if nstyle in ('Heading 1', 'Heading 2'):
                            break  # next section starts

                        if nstyle == 'Heading 3' and ntext:
                            if current_subpoint is not None:
                                subpoints.append(current_subpoint)
                            current_subpoint = {'title': ntext, 'bullets': []}
                            i += 1
                            continue

                        if nstyle == 'List Paragraph' and ntext:
                            if current_subpoint is not None:
                                current_subpoint['bullets'].append(ntext)
                            else:
                                collected_bullets.append(ntext)
                            i += 1
                            continue

                        # Body Text / Normal / blank - descriptive filler, skip
                        i += 1
                        continue

                if current_subpoint is not None:
                    subpoints.append(current_subpoint)

                if current_h2 not in sections:
                    if subpoints:
                        sections[current_h2] = {'type': 'subpoints', 'data': subpoints}
                    else:
                        sections[current_h2] = {'type': 'list', 'data': collected_bullets}
                continue

        i += 1

    return {
        'sections': sections,
        'section_order': section_order,
    }


def _find_section(parsed, *name_fragments):
    """Case-insensitive lookup of a section whose heading contains all fragments."""
    for heading, content in parsed['sections'].items():
        h = heading.lower()
        if all(f.lower() in h for f in name_fragments):
            return content
    return None


def to_affidavit_context(parsed):
    """Flatten the generic parsed sections into the specific fields the
    affidavit drafting/rendering pipeline needs. Raises if a required
    section is missing, so failures are explicit rather than silent."""
    case = _find_section(parsed, 'court') or _find_section(parsed, 'case')
    deponent = _find_section(parsed, 'deponent')
    points = _find_section(parsed, 'reply', 'point') or _find_section(parsed, 'point')
    prayer = _find_section(parsed, 'prayer')
    attestation = _find_section(parsed, 'attestation')
    advocate = _find_section(parsed, 'advocate')

    missing = [n for n, v in [
        ('Court and Case Details', case), ('Deponent Details', deponent),
        ('Reply Points', points), ('Prayer', prayer),
        ('Attestation Details', attestation), ('Advocate', advocate),
    ] if v is None]
    if missing:
        raise ValueError(f"Case Information doc is missing expected section(s): {', '.join(missing)}")

    return {
        'case': case['data'],
        'deponent': deponent['data'],
        'reply_points': points['data'],  # list of {'title', 'bullets'}
        'prayer_items': prayer['data'],  # list of strings
        'attestation': attestation['data'],
        'advocate': advocate['data'],
    }


if __name__ == '__main__':
    import sys, json
    parsed = parse_case_info(sys.argv[1])
    ctx = to_affidavit_context(parsed)
    print(json.dumps(ctx, indent=2, ensure_ascii=False))
