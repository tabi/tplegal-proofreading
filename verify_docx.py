#!/usr/bin/env python3
"""Verify DOCX integrity after correction pipeline.

Compares the original document with a corrected version to detect text loss,
missing paragraphs, or structural degradation.

Usage:
    python verify_docx.py original.docx corrected.docx [--strict] [--dump]

Checks performed:
    1. Paragraph count comparison
    2. Total character count comparison (warns if >1% loss)
    3. Per-paragraph text diff (flags missing/truncated paragraphs)
    4. Structural element count (tables, images, headers/footers)

Exit codes:
    0 = OK (no significant differences)
    1 = WARNING (minor differences detected)
    2 = ERROR (significant text loss or missing paragraphs)
    3 = FATAL (file cannot be read)
"""

import argparse
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import defusedxml.minidom as minidom_mod

from minidom_helpers import find_elements as _find_elements, extract_paragraph_text as _extract_paragraph_text


@dataclass
class DocStats:
    paragraphs: list[str] = field(default_factory=list)
    total_chars: int = 0
    table_count: int = 0
    image_count: int = 0
    header_footer_count: int = 0


def extract_stats(docx_path: str, mode: str = "visible") -> DocStats:
    """Extract text and structural stats from a DOCX file.

    Args:
        docx_path: Path to the DOCX file.
        mode: 'visible' — text as Word displays it.
              'original' — text before tracked changes were applied.
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")

    stats = DocStats()

    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()

        # Count structural elements
        stats.header_footer_count = sum(
            1 for n in names
            if n.startswith("word/header") or n.startswith("word/footer")
        )

        # Parse document.xml
        if "word/document.xml" not in names:
            raise ValueError(f"No word/document.xml in {docx_path}")

        doc_xml = zf.read("word/document.xml").decode("utf-8")
        dom = minidom_mod.parseString(doc_xml)
        root = dom.documentElement

        # Extract paragraphs
        for p_elem in _find_elements(root, "p"):
            text = _extract_paragraph_text(p_elem, mode=mode)
            stats.paragraphs.append(text)
            stats.total_chars += len(text)

        # Count tables
        stats.table_count = len(_find_elements(root, "tbl"))

        # Count images (drawings + legacy VML)
        stats.image_count = (
            len(_find_elements(root, "drawing"))
            + len(_find_elements(root, "pict"))
        )

    return stats


def compare(original_path: str, corrected_path: str, strict: bool = False, dump: bool = False, quiet: bool = False) -> int:
    """Compare original and corrected DOCX files.

    Returns exit code: 0=OK, 1=WARNING, 2=ERROR, 3=FATAL
    """
    try:
        orig = extract_stats(original_path, mode="visible")
    except Exception as e:
        print(f"FATAL: Cannot read original file: {e}", file=sys.stderr)
        return 3

    try:
        corr = extract_stats(corrected_path, mode="original")
    except Exception as e:
        print(f"FATAL: Cannot read corrected file: {e}", file=sys.stderr)
        return 3

    exit_code = 0
    issues = []

    # --- 1. Paragraph count ---
    orig_p = len(orig.paragraphs)
    corr_p = len(corr.paragraphs)
    if orig_p != corr_p:
        diff = orig_p - corr_p
        msg = f"Paragraph count: original={orig_p}, corrected={corr_p} (diff={diff:+d})"
        if abs(diff) > 2:
            issues.append(("ERROR", msg))
            exit_code = max(exit_code, 2)
        else:
            issues.append(("WARN", msg))
            exit_code = max(exit_code, 1)
    else:
        issues.append(("OK", f"Paragraph count: {orig_p} = {corr_p}"))

    # --- 2. Total character count ---
    if orig.total_chars > 0:
        loss = orig.total_chars - corr.total_chars
        loss_pct = (loss / orig.total_chars) * 100
        msg = f"Characters: original={orig.total_chars}, corrected={corr.total_chars} (loss={loss:+d}, {loss_pct:+.1f}%)"

        if loss_pct > 5.0:
            issues.append(("ERROR", msg))
            exit_code = max(exit_code, 2)
        elif loss_pct > 1.0:
            issues.append(("WARN", msg))
            exit_code = max(exit_code, 1)
        else:
            issues.append(("OK", msg))
    else:
        issues.append(("WARN", "Original document has 0 characters"))

    # --- 3. Per-paragraph diff (only non-empty paragraphs) ---
    orig_nonempty = [(i, p) for i, p in enumerate(orig.paragraphs) if p.strip()]
    corr_nonempty = [(i, p) for i, p in enumerate(corr.paragraphs) if p.strip()]

    # Build text→index map for corrected paragraphs
    corr_text_set = {p for _, p in corr_nonempty}

    missing_paragraphs = []
    truncated_paragraphs = []

    for orig_idx, orig_text in orig_nonempty:
        if orig_text in corr_text_set:
            continue
        # Check for truncation — is a prefix present?
        found_truncated = False
        for _, corr_text in corr_nonempty:
            if len(corr_text) > 20 and orig_text.startswith(corr_text):
                lost = len(orig_text) - len(corr_text)
                truncated_paragraphs.append((orig_idx, lost, orig_text[:80]))
                found_truncated = True
                break
            if len(orig_text) > 20 and corr_text.startswith(orig_text[:20]):
                # Partial match — likely the same paragraph with edits (OK)
                found_truncated = True
                break
        if not found_truncated:
            # Check if it's genuinely missing or just edited
            # Use a rough similarity check: first 30 chars match
            prefix = orig_text[:30] if len(orig_text) >= 30 else orig_text
            has_rough_match = any(
                corr_text.startswith(prefix) or prefix in corr_text
                for _, corr_text in corr_nonempty
            )
            if not has_rough_match and len(orig_text) > 10:
                missing_paragraphs.append((orig_idx, orig_text[:100]))

    if missing_paragraphs:
        issues.append(("ERROR", f"Missing paragraphs: {len(missing_paragraphs)}"))
        for idx, preview in missing_paragraphs[:5]:
            issues.append(("  ", f"  ¶{idx}: \"{preview}...\""))
        if len(missing_paragraphs) > 5:
            issues.append(("  ", f"  ... and {len(missing_paragraphs) - 5} more"))
        exit_code = max(exit_code, 2)

    if truncated_paragraphs:
        issues.append(("WARN", f"Truncated paragraphs: {len(truncated_paragraphs)}"))
        for idx, lost, preview in truncated_paragraphs[:5]:
            issues.append(("  ", f"  ¶{idx}: lost {lost} chars — \"{preview}...\""))
        exit_code = max(exit_code, 2)

    if not missing_paragraphs and not truncated_paragraphs:
        issues.append(("OK", "No missing or truncated paragraphs detected"))

    # --- 4. Structural elements ---
    if orig.table_count != corr.table_count:
        issues.append(("WARN", f"Tables: original={orig.table_count}, corrected={corr.table_count}"))
        exit_code = max(exit_code, 1)

    if orig.image_count != corr.image_count:
        issues.append(("WARN", f"Images: original={orig.image_count}, corrected={corr.image_count}"))
        exit_code = max(exit_code, 1)

    # --- Output ---
    status_label = {0: "PASS", 1: "WARNING", 2: "FAIL", 3: "FATAL"}

    if not quiet:
        print(f"\n{'='*60}")
        print(f"DOCX INTEGRITY CHECK: {status_label[exit_code]}")
        print(f"{'='*60}")
        print(f"Original:  {original_path}")
        print(f"Corrected: {corrected_path}")
        print(f"{'-'*60}")

        for level, msg in issues:
            if level == "  ":
                print(msg)
            else:
                print(f"[{level:>5}] {msg}")

        print(f"{'='*60}\n")

    # --- Optional: dump paragraphs ---
    if dump and not quiet:
        print("--- ORIGINAL PARAGRAPHS (non-empty) ---")
        for i, p in orig_nonempty:
            print(f"  ¶{i:03d} ({len(p):>5} chars): {p[:120]}")
        print(f"\n--- CORRECTED PARAGRAPHS (non-empty) ---")
        for i, p in corr_nonempty:
            print(f"  ¶{i:03d} ({len(p):>5} chars): {p[:120]}")
        print()

    return exit_code


def main():
    parser = argparse.ArgumentParser(
        description="Verify DOCX integrity after correction pipeline"
    )
    parser.add_argument("original", help="Original DOCX file")
    parser.add_argument("corrected", help="Corrected DOCX file")
    parser.add_argument(
        "--strict", action="store_true",
        help="Treat warnings as errors (exit code 2)"
    )
    parser.add_argument(
        "--dump", action="store_true",
        help="Dump all non-empty paragraphs for manual comparison"
    )
    args = parser.parse_args()

    exit_code = compare(args.original, args.corrected, strict=args.strict, dump=args.dump)

    if args.strict and exit_code == 1:
        exit_code = 2

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
