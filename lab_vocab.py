#!/usr/bin/env python3
"""
Caption / comment vocabulary for the Lab Report Reviewer.

Pure regex + lookup helpers (no openpyxl / PIL) for reading what a micrograph
caption says: its picture number, etch status, etchant and heat-treatment
condition, plus the alloy-name pattern. Kept separate from the parsing and the
rules so the vocabulary lives in one place. Imported (and re-exported) by
``lab_review`` — external callers can keep importing these from there.
"""
import re

# Caption / comment integrity vocabulary.
_PICNUM   = re.compile(r'picture\s*(\d+)', re.I)
_ETCH_PAT = re.compile(r'etch|unetched|as[-\s]?polished|kalling|glyceregia|oxalic|'
                       r'marble|nital|vilella|murakami|aqua\s*regia|electrolytic', re.I)
# Captions that explicitly state the section was NOT etched (vs simply omitting
# the etch status). 'etch' in _ETCH_PAT also matches 'unetched', so these are
# recognised as *having* an etch status — they just need surfacing on their own.
_UNETCHED_PAT = re.compile(r'\bunetched\b|as[-\s]?polished', re.I)
_ALLOY_PAT = re.compile(
    r'\b(?:IN[-\s]?\d{3}(?:LC)?|GTD[-\s]?\d{3}|Ren[eé][-\s]?\d+|Nimonic[-\s]?\d+|'
    r'Inconel[-\s]?\d+|Hastelloy[-\s]?\w?|Waspaloy|Mar[-\s]?M[-\s]?\d+|'
    r'FSX[-\s]?\d+|Udimet[-\s]?\d+|C[-\s]?263)\b', re.I)


def _norm_alloy(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


# Etchant vocabulary (ordered: multi-word / specific first, generic last).
_ETCHANT_VOCAB = [
    (r'unetched|as[-\s]?polished', 'Unetched'),
    (r'waterless\s*kalling',       'Waterless Kalling'),
    (r'\bkalling',                 'Kalling'),
    (r'oxalic',                    'Oxalic Acid'),
    (r'glyceregia',                'Glyceregia'),
    (r'\bmarble',                  "Marble's"),
    (r'\bnital\b',                 'Nital'),
    (r'vilella',                   "Vilella's"),
    (r'murakami',                  'Murakami'),
    (r'aqua\s*regia',              'Aqua Regia'),
    (r'electrolytic',              'Electrolytic'),
    (r'\betch',                    'Etched (unspecified)'),
]


def caption_etchant(text):
    """Canonical etchant named in a caption, or None."""
    t = text or ''
    for pat, name in _ETCHANT_VOCAB:
        if re.search(pat, t, re.I):
            return name
    return None


def report_etchants(pictures):
    """(magnification→etchant map, primary named etchant) from a report's captions."""
    by_mag, counts = {}, {}
    for label, cap in pictures or []:
        text = f"{label} {cap or ''}"
        et = caption_etchant(text)
        if et and et not in ('Unetched', 'Etched (unspecified)'):
            counts[et] = counts.get(et, 0) + 1
        if et:
            for m in re.finditer(r'(\d{2,4})\s*[xX]\b', text):
                by_mag.setdefault(f"{m.group(1)}x", et)
    primary = max(counts, key=counts.get) if counts else None
    if primary is None and by_mag:        # no named etchant → most common caption etchant
        vals = list(by_mag.values())
        primary = max(set(vals), key=vals.count)
    return by_mag, primary


def image_etchant(image_mag, by_mag, primary):
    """Best-effort etchant for one micrograph (caption etchant for its magnification)."""
    if image_mag and image_mag in by_mag:
        return by_mag[image_mag]
    return primary or 'Unspecified'


# Heat-treatment condition vocabulary (ordered by repair sequence; specific
# first). The condition usually varies per picture within a report.
_HT_VOCAB = [
    (r'post[-\s]*ag(?:e?ing|ed)|after\s*ag(?:e?ing)|\bre[-\s]*ag|\baged\b|\bage?ing\b', 'Post-ageing'),
    (r'stress[-\s]*relief|after\s*stress|\bSR\b\s*HT|post[-\s]*weld', 'Post stress-relief'),
    # As-received / pre-solution before post-solution so "Pre-Solution" isn't
    # caught by the post-solution "re-solution" alternative.
    (r'pre[-\s]*solution|as[-\s]*received|as\s*received|receiving|incoming|as[-\s]*is|service[-\s]*exposed', 'As-received'),
    (r'post[-\s]*solution|after\s*solution|solution(?:ed|ing)?\s*(?:ht|treat)|\bre[-\s]*solution', 'Post-solution'),
]

# Display order for HT groups (process sequence).
HT_ORDER = ['As-received', 'Post-solution', 'Post stress-relief', 'Post-ageing', 'Unspecified']


def caption_ht(text):
    """Canonical heat-treatment condition named in a caption, or None."""
    t = text or ''
    for pat, name in _HT_VOCAB:
        if re.search(pat, t, re.I):
            return name
    return None


def report_ht(pictures):
    """(magnification→HT map, primary HT) from a report's captions.

    HT typically varies per picture (pre/post-solution, stress relief, ageing),
    so the per-magnification map is the main output.
    """
    by_mag, counts = {}, {}
    for label, cap in pictures or []:
        text = f"{label} {cap or ''}"
        ht = caption_ht(text)
        if ht:
            counts[ht] = counts.get(ht, 0) + 1
            for m in re.finditer(r'(\d{2,4})\s*[xX]\b', text):
                by_mag.setdefault(f"{m.group(1)}x", ht)
    primary = max(counts, key=counts.get) if counts else None
    return by_mag, primary


def image_ht(image_mag, by_mag, primary):
    """Best-effort HT condition for one micrograph (caption HT for its magnification)."""
    if image_mag and image_mag in by_mag:
        return by_mag[image_mag]
    return primary or 'Unspecified'
