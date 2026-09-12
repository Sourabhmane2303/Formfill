"""
Generic fuzzy field mapper shared by both template engines:
  - content-control engine: candidates are {control_index: [tag, title]}
  - Jinja/docxtpl engine:    candidates are {variable_name: [variable_name]}

Matches extracted (label, value) pairs from an info doc onto whichever
candidate has the closest-matching alias string.
"""
from rapidfuzz import fuzz

SYNONYMS = {
    'name': ['full name', 'employee name', 'candidate name', 'petitioner name'],
    'dob': ['date of birth', 'birth date'],
    'address': ['current address', 'residential address', 'home address'],
    'email': ['email address', 'e-mail', 'e-mail address'],
    'phone': ['phone number', 'contact number', 'mobile number', 'mobile'],
    'designation': ['job title', 'role', 'position'],
    'joining date': ['date of joining', 'doj'],
    'petition no': ['petition number', 'case no', 'case number']
}


def _normalize(s):
    return (s or '').strip().lower().replace('_', ' ').replace('-', ' ')


def _expand_synonyms(label_norm):
    variants = {label_norm}
    for canonical, alts in SYNONYMS.items():
        group = {canonical, *alts}
        if label_norm in group:
            variants |= group
    return variants


def map_fields(extracted_fields, candidates, score_threshold=60):
    """
    extracted_fields: list of (label, value)
    candidates: dict {candidate_id: [alias_string, ...]}

    Returns: mapping {candidate_id: value}, matched_report (list of dicts),
             unmatched (list of (label, value)), unfilled_ids (list)
    """
    flat_candidates = [
        (cid, _normalize(alias))
        for cid, aliases in candidates.items()
        for alias in aliases if alias
    ]

    mapping = {}
    matched_report = []
    unmatched = []
    used = set()

    for label, value in extracted_fields:
        label_norm = _normalize(label)
        variants = sorted(_expand_synonyms(label_norm))

        best_id, best_score = None, 0
        for variant in variants:
            for cid, cand_norm in flat_candidates:
                if cid in used:
                    continue
                score = fuzz.token_sort_ratio(variant, cand_norm)
                if score > best_score:
                    best_score, best_id = score, cid

        if best_id is not None and best_score >= score_threshold:
            used.add(best_id)
            mapping[best_id] = value
            matched_report.append({
                'label': label, 'value': value,
                'matched_field': best_id, 'score': best_score,
            })
        else:
            unmatched.append((label, value))

    unfilled_ids = [cid for cid in candidates if cid not in mapping]
    return mapping, matched_report, unmatched, unfilled_ids
