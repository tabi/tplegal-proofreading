#!/usr/bin/env python3
"""
Apply proofreading corrections to an unpacked .docx as tracked changes.

Input:  unpacked docx directory + JSON corrections file
Output: modified document.xml with w:del/w:ins tracked changes

Corrections JSON format:
[
  {
    "original": "tekst z bledem",
    "corrected": "tekst z błędem",
    "note": "ortografia"           # optional - added as comment
  },
  ...
]

Usage:
  python apply_corrections.py unpacked/ corrections.json [--author "Korektor"] [--date "2026-03-18T00:00:00Z"]
"""

import argparse
import json
import logging
import os
from datetime import datetime, timezone

from ooxml import (
    W,
    IdCounter,
    find_max_id,
    get_paragraph_runs,
    get_paragraph_text,
    get_paragraph_text_from_runs,
    get_rpr,
    make_del,
    make_ins,
    make_run,
)

log = logging.getLogger(__name__)


def apply_correction(para, original, corrected, author, date_str, id_counter):
    """Apply a single correction to a paragraph as tracked changes.

    Accepts an IdCounter instance. Returns True if the correction was applied.
    """
    runs = get_paragraph_runs(para)
    if not runs:
        return False

    full_text = get_paragraph_text_from_runs(runs)
    idx = full_text.find(original)
    if idx == -1:
        idx_lower = full_text.lower().find(original.lower())
        if idx_lower != -1:
            idx = idx_lower
            original = full_text[idx:idx + len(original)]
        else:
            log.warning("'%s' not found in paragraph text", original)
            return False

    orig_start = idx
    orig_end = idx + len(original)

    affected = []
    for run in runs:
        if run['start'] < orig_end and run['end'] > orig_start:
            affected.append(run)

    if not affected:
        return False

    first_rpr = affected[0]['rpr']

    elements_to_insert = []

    # Prefix
    first_run = affected[0]
    if first_run['start'] < orig_start:
        prefix_text = first_run['text'][:orig_start - first_run['start']]
        elements_to_insert.append(make_run(prefix_text, first_rpr))

    # w:del — multi-run aware: each affected run contributes its slice
    del_elem = make_del(
        '', first_rpr,
        id_counter=id_counter, author=author, date_str=date_str,
    )
    # Remove the empty placeholder run that make_del created
    for child in list(del_elem):
        del_elem.remove(child)
    for run in affected:
        slice_start = max(0, orig_start - run['start'])
        slice_end = min(len(run['text']), orig_end - run['start'])
        del_text = run['text'][slice_start:slice_end]
        if del_text:
            del_elem.append(make_run(del_text, run['rpr'], is_delete=True))
    elements_to_insert.append(del_elem)

    # w:ins
    elements_to_insert.append(make_ins(
        corrected, first_rpr,
        id_counter=id_counter, author=author, date_str=date_str,
    ))

    # Suffix
    last_run = affected[-1]
    if last_run['end'] > orig_end:
        suffix_text = last_run['text'][orig_end - last_run['start']:]
        elements_to_insert.append(make_run(suffix_text, last_run['rpr']))

    # Replace affected runs in DOM
    first_elem = affected[0]['elem']
    parent = first_elem.getparent()
    insert_pos = list(parent).index(first_elem)

    for run in affected:
        if run['elem'].getparent() is parent:
            parent.remove(run['elem'])

    for i, elem in enumerate(elements_to_insert):
        parent.insert(insert_pos + i, elem)

    return True


def main():
    parser = argparse.ArgumentParser(description='Apply corrections as tracked changes')
    parser.add_argument('unpacked_dir', help='Path to unpacked docx directory')
    parser.add_argument('corrections_json', help='Path to corrections JSON file')
    parser.add_argument('--author', default='Korektor AI', help='Author name for tracked changes')
    parser.add_argument('--date', default=None, help='Date for tracked changes (ISO format)')
    args = parser.parse_args()

    if args.date is None:
        args.date = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

    with open(args.corrections_json, 'r', encoding='utf-8') as f:
        corrections = json.load(f)

    log.info("Loaded %d corrections", len(corrections))

    doc_path = os.path.join(args.unpacked_dir, 'word', 'document.xml')
    tree = etree.parse(doc_path)
    root = tree.getroot()

    id_counter = IdCounter(find_max_id(tree))

    paragraphs = root.findall(f'.//{W}p')

    applied = 0
    for corr in corrections:
        original = corr['original']
        corrected = corr['corrected']

        if original == corrected:
            continue

        found = False
        for para in paragraphs:
            full_text = get_paragraph_text(para)
            if original in full_text or original.lower() in full_text.lower():
                if apply_correction(para, original, corrected, args.author, args.date, id_counter):
                    applied += 1
                    found = True
                    log.info("Applied: '%s' → '%s'", original, corrected)
                    break

        if not found:
            log.warning("SKIPPED: '%s' not found in document", original)

    log.info("Applied %d/%d corrections", applied, len(corrections))

    tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    log.info("Saved to %s", doc_path)


if __name__ == '__main__':
    logging.basicConfig(
        level=os.environ.get('LOG_LEVEL', 'INFO').upper(),
        format='%(levelname)s: %(message)s',
    )
    main()
