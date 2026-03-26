#!/usr/bin/env python3
"""
Proofreader for DOCX files — applies corrections as tracked changes.
Uses pluggable text checkers (LanguageTool API by default).

Usage:
    python proofread_docx.py <input.docx> [output.docx]

The output file will contain tracked changes (w:del/w:ins) for each correction.
Formatting is preserved — only text content is modified.
"""

import argparse
import logging
import os
import shutil
import sys
import tempfile
import time

from lxml import etree
from datetime import datetime, timezone

from ooxml import (
    W,
    IdCounter,
    find_max_id,
    get_paragraph_runs,
    get_paragraph_text_from_runs,
    make_run,
    make_del,
    make_ins,
)
from docx_io import unpack, pack
from checkers import LanguageToolChecker, CompositeChecker, polish_legal_rules
from verify_docx import compare

log = logging.getLogger(__name__)

# ── Configuration (overridable via env vars) ───────────────────────────────

AUTHOR = os.environ.get('PROOFREAD_AUTHOR', 'Korektor')


# ── Core: apply a single match to a paragraph ─────────────────────────────

def _apply_single_match(p_elem, match, id_counter, author, date_str):
    """Apply one match as a tracked change. Re-scans runs each time.

    Returns True if the correction was applied, False otherwise.
    """
    m_start = match['offset']
    m_end = match['end']
    replacement = match['replacement']

    runs = get_paragraph_runs(p_elem)
    if not runs:
        return False

    full_text = get_paragraph_text_from_runs(runs)

    # Verify the match still refers to the expected text
    if m_end > len(full_text):
        return False

    # Find affected runs
    affected = []
    for i, run in enumerate(runs):
        if run['end'] > m_start and run['start'] < m_end:
            affected.append((i, run))

    if not affected:
        return False

    _, first_run = affected[0]
    _, last_run = affected[-1]
    rpr = first_run['rpr']

    # Calculate prefix/suffix within the affected runs
    prefix_text = first_run['text'][:m_start - first_run['start']]
    suffix_text = last_run['text'][m_end - last_run['start']:]
    deleted_text = full_text[m_start:m_end]

    # Build replacement elements
    new_elements = []
    if prefix_text:
        new_elements.append(make_run(prefix_text, first_run['rpr']))
    new_elements.append(make_del(deleted_text, rpr, id_counter=id_counter, author=author, date_str=date_str))
    new_elements.append(make_ins(replacement, rpr, id_counter=id_counter, author=author, date_str=date_str))
    if suffix_text:
        new_elements.append(make_run(suffix_text, last_run['rpr']))

    # Find insertion point
    first_elem = first_run['elem']
    parent = first_elem.getparent()
    if parent is None:
        return False

    insert_idx = list(parent).index(first_elem)

    # Remove affected run elements
    for _, run in affected:
        elem = run['elem']
        if elem.getparent() is not None:
            elem.getparent().remove(elem)

    # Insert new elements
    for j, new_elem in enumerate(new_elements):
        parent.insert(insert_idx + j, new_elem)

    return True


def apply_corrections_to_paragraph(p_elem, matches, id_counter, author, date_str):
    """Apply all matches to a paragraph as tracked changes.

    Processes matches right-to-left so that offsets of earlier matches stay valid.
    Re-scans runs for each match to handle single-run paragraphs correctly.
    """
    if not matches:
        return 0

    changes = 0
    for match in reversed(matches):
        if _apply_single_match(p_elem, match, id_counter, author, date_str):
            changes += 1

    return changes


# ── Main processing pipeline ──────────────────────────────────────────────

def process_document(unpacked_dir, checker):
    """Process document.xml: check each paragraph and apply corrections."""
    doc_path = os.path.join(unpacked_dir, 'word', 'document.xml')

    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(doc_path, parser)
    root = tree.getroot()

    id_counter = IdCounter(find_max_id(tree) + 100)
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    paragraphs = list(root.iter(f'{W}p'))

    total_changes = 0
    para_count = 0

    for p_elem in paragraphs:
        runs = get_paragraph_runs(p_elem)
        text = get_paragraph_text_from_runs(runs)

        if not text.strip() or len(text.strip()) < 3:
            continue

        para_count += 1
        matches = checker.check(text)

        if matches:
            n = apply_corrections_to_paragraph(p_elem, matches, id_counter, AUTHOR, date_str)
            total_changes += n

            if n > 0:
                for m in matches:
                    old = text[m['offset']:m['end']]
                    log.info('[%s] "%s" -> "%s"  (%s)', m['category'], old, m['replacement'], m['message'])

        if hasattr(checker, 'request_delay'):
            time.sleep(checker.request_delay)

    log.info("Processed %d paragraphs, applied %d corrections.", para_count, total_changes)

    tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return total_changes


def build_checker():
    """Build the default checker pipeline."""
    composite = CompositeChecker()
    composite.add(LanguageToolChecker())
    composite.add(polish_legal_rules())
    return composite


def main():
    parser = argparse.ArgumentParser(
        description='Proofread a DOCX file and produce tracked changes.'
    )
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('output', nargs='?', help='Output .docx file (default: input_proofread.docx)')
    parser.add_argument('--quiet', '-q', action='store_true', help='Summary only — suppress per-correction logs')
    parser.add_argument('--no-verify', action='store_true', help='Skip integrity check after correction')
    args = parser.parse_args()

    log_level = logging.WARNING if args.quiet else os.environ.get('LOG_LEVEL', 'INFO').upper()
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    input_path = args.input
    output_path = args.output or input_path.replace('.docx', '_proofread.docx')

    if not os.path.exists(input_path):
        print(f"Error: file not found: {input_path}", file=sys.stderr)  # noqa: T201
        sys.exit(1)

    log.info("Input:  %s", input_path)
    log.info("Output: %s", output_path)

    checker = build_checker()

    unpacked_dir = tempfile.mkdtemp(prefix='proofread_')
    try:
        log.info("[1/4] Unpacking DOCX...")
        unpack(input_path, unpacked_dir)

        log.info("[2/4] Checking text...")
        n_changes = process_document(unpacked_dir, checker)

        if n_changes == 0:
            log.info("No corrections found — document looks clean.")

        log.info("[3/4] Repacking DOCX...")
        pack(unpacked_dir, output_path, original_docx=input_path)
    finally:
        shutil.rmtree(unpacked_dir, ignore_errors=True)

    if not args.no_verify:
        log.info("[4/4] Verifying integrity...")
        verify_code = compare(input_path, output_path, quiet=args.quiet)
        if verify_code >= 2:
            if args.quiet:
                print(f"\u26a0 Integrity FAILED — possible text loss in {output_path}", file=sys.stderr)  # noqa: T201
            else:
                log.error("Integrity check FAILED — possible text loss in %s", output_path)
            sys.exit(2)
        elif verify_code == 1:
            if args.quiet:
                print(f"\u26a0 Integrity WARNING — review {output_path} manually.")  # noqa: T201
            else:
                log.warning("Integrity check passed with warnings — review output manually.")
        else:
            if args.quiet:
                pass  # OK — no output needed for integrity in quiet+pass
    else:
        log.info("[4/4] Skipping integrity check (--no-verify).")

    if args.quiet:
        if n_changes > 0:
            print(f"\u2713 {n_changes} corrections applied to {output_path}")  # noqa: T201
        else:
            print("\u2713 No corrections needed \u2014 document is clean.")  # noqa: T201
    else:
        log.info("Done! Output: %s", output_path)

    return 0


if __name__ == '__main__':
    main()
