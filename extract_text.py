#!/usr/bin/env python3
"""Extract plain text from a DOCX file for proofreading.

Reads word/document.xml body paragraphs and outputs numbered text.
Skips empty paragraphs, field codes, and existing tracked changes.

Usage:
    extract-text input.docx
    extract-text input.docx > tekst.txt
"""

import argparse
import sys
import zipfile
from pathlib import Path

import defusedxml.minidom as minidom_mod

from minidom_helpers import extract_paragraph_text, find_elements, match_local


def _is_field_paragraph(p_elem) -> bool:
    """Check if paragraph contains only field code elements (no real text)."""
    has_field = False
    has_text = False
    for child in p_elem.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
        name = child.localName or child.tagName
        if match_local(name, "r"):
            for sub in child.childNodes:
                if sub.nodeType != sub.ELEMENT_NODE:
                    continue
                sname = sub.localName or sub.tagName
                if match_local(sname, "fldChar") or match_local(sname, "instrText"):
                    has_field = True
                elif match_local(sname, "t"):
                    if sub.firstChild and hasattr(sub.firstChild, "data") and sub.firstChild.data.strip():
                        has_text = True
    return has_field and not has_text


def extract_text(docx_path: str) -> list[str]:
    """Extract numbered paragraph texts from a DOCX file.

    Returns list of strings like '001: Tekst akapitu...'
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")

    with zipfile.ZipFile(path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            raise ValueError(f"No word/document.xml in {docx_path}")
        doc_xml = zf.read("word/document.xml").decode("utf-8")

    dom = minidom_mod.parseString(doc_xml)
    root = dom.documentElement

    # Find body element
    body = None
    for child in root.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            name = child.localName or child.tagName
            if match_local(name, "body"):
                body = child
                break
    if body is None:
        body = root

    lines = []
    num = 0
    for p_elem in find_elements(body, "p"):
        # Skip paragraphs inside tracked changes at document level
        # (these are rare but possible in some OOXML constructs)

        # Skip field-only paragraphs
        if _is_field_paragraph(p_elem):
            continue

        text = extract_paragraph_text(p_elem, mode="visible")

        # Skip empty paragraphs
        if not text.strip():
            continue

        num += 1
        lines.append(f"\u00b6{num:03d}: {text}")

    return lines


def main():
    parser = argparse.ArgumentParser(
        description="Extract plain text from a DOCX file for proofreading"
    )
    parser.add_argument("docx", help="Input DOCX file")
    args = parser.parse_args()

    try:
        lines = extract_text(args.docx)
    except (FileNotFoundError, ValueError, zipfile.BadZipFile) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(2)

    for line in lines:
        print(line)


if __name__ == "__main__":
    main()
