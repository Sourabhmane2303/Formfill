"""
Utilities for reading and writing Word "content controls" (w:sdt elements).

python-docx does not expose a high-level API for content control values,
so we operate on the underlying XML tree directly (docx.oxml / lxml).

Supports plain-text content controls (<w:sdt> with <w:text/> in sdtPr).
Rich-text / dropdown / date-picker controls are detected but only their
plain text run content is read/written (sufficient for text-fill use
cases; dropdown *selection* is out of scope here).
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class ContentControl:
    def __init__(self, sdt_element):
        self.sdt = sdt_element

    @property
    def tag(self):
        sdtPr = self.sdt.find(qn('w:sdtPr'))
        if sdtPr is None:
            return None
        tag_el = sdtPr.find(qn('w:tag'))
        return tag_el.get(qn('w:val')) if tag_el is not None else None

    @property
    def title(self):
        sdtPr = self.sdt.find(qn('w:sdtPr'))
        if sdtPr is None:
            return None
        alias_el = sdtPr.find(qn('w:alias'))
        return alias_el.get(qn('w:val')) if alias_el is not None else None

    @property
    def text(self):
        sdtContent = self.sdt.find(qn('w:sdtContent'))
        if sdtContent is None:
            return ''
        texts = sdtContent.findall('.//' + qn('w:t'))
        return ''.join(t.text or '' for t in texts)

    def set_text(self, new_text):
        """Replace the content control's text with new_text, collapsing
        it into a single run so formatting stays predictable."""
        sdtContent = self.sdt.find(qn('w:sdtContent'))
        if sdtContent is None:
            return False

        # Try to preserve run formatting (rPr) from the first existing run
        existing_runs = sdtContent.findall(qn('w:r'))
        rPr = None
        if existing_runs:
            first_rPr = existing_runs[0].find(qn('w:rPr'))
            if first_rPr is not None:
                rPr = first_rPr

        # Remove all existing child runs/paragraphs-with-runs
        for child in list(sdtContent):
            sdtContent.remove(child)

        r = OxmlElement('w:r')
        if rPr is not None:
            import copy
            r.append(copy.deepcopy(rPr))
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = new_text
        r.append(t)
        sdtContent.append(r)
        return True

    def __repr__(self):
        return f"<ContentControl tag={self.tag!r} title={self.title!r} text={self.text!r}>"


def iter_content_controls(document):
    """Yield ContentControl objects for every w:sdt in the document body
    (including ones nested inside tables)."""
    body = document.element.body
    for sdt in body.iter(qn('w:sdt')):
        yield ContentControl(sdt)


def get_schema(document):
    """Return a list of dicts describing each content control:
    tag, title, current text. This is the reusable 'template schema'."""
    schema = []
    for cc in iter_content_controls(document):
        schema.append({
            'tag': cc.tag,
            'title': cc.title,
            'current_text': cc.text,
        })
    return schema
