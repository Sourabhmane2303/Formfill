"""
Extracts fields from a free-text 'info doc'.

Supports two shapes:
  1. Single-line:  "Label: value"
  2. Block fields: "Label:" on its own line, followed by multiple lines
     of content, terminated by a line "End <Label>" (case-insensitive).
     Used for content that spans several lines/paragraphs, e.g. the
     numbered body paragraphs of an affidavit.

Reads paragraphs from a .docx in document order (tables not handled
here since legal info docs are typically plain running text).
"""
import re
from docx import Document

LINE_PATTERN = re.compile(r'^\s*(?P<label>[^:]{1,60}):\s*(?P<value>.*\S)\s*$')
LABEL_ONLY_PATTERN = re.compile(r'^\s*(?P<label>[^:]{1,60}):\s*$')
END_PATTERN = re.compile(r'^\s*end\s+(?P<label>.+?)\s*$', re.IGNORECASE)


def extract_fields(docx_path):
    """Returns an ordered list of (label, value) tuples, where block
    fields have their inner lines joined with '\\n'."""
    doc = Document(docx_path)
    lines = [p.text for p in doc.paragraphs]  # keep blank lines as boundaries

    results = []
    i = 0
    n = len(lines)
    while i < n:
        raw = lines[i]
        stripped = raw.strip()
        if not stripped:
            i += 1
            continue

        label_only_match = LABEL_ONLY_PATTERN.match(stripped)
        if label_only_match:
            label = label_only_match.group('label').strip()
            block_lines = []
            j = i + 1
            while j < n:
                candidate = lines[j].strip()
                end_match = END_PATTERN.match(candidate)
                if end_match and end_match.group('label').strip().lower() == label.lower():
                    break
                block_lines.append(lines[j])
                j += 1
            # Trim leading/trailing blank lines in the captured block
            while block_lines and not block_lines[0].strip():
                block_lines.pop(0)
            while block_lines and not block_lines[-1].strip():
                block_lines.pop()
            results.append((label, "\n".join(block_lines)))
            i = j + 1  # skip past the 'End <label>' line
            continue

        line_match = LINE_PATTERN.match(stripped)
        if line_match:
            results.append((line_match.group('label').strip(), line_match.group('value').strip()))
        i += 1

    return results
