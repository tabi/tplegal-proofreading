"""DOCX pack/unpack — self-contained, no external script dependencies.

A .docx file is a ZIP archive with a specific structure. These functions
handle unpacking and repacking while preserving the original archive's
metadata (content types, relationships, etc.).
"""

from __future__ import annotations

import logging
import os
import shutil
import zipfile

log = logging.getLogger(__name__)


def unpack(docx_path: str, output_dir: str) -> None:
    """Unpack a .docx file into a directory.

    Args:
        docx_path: Path to the .docx file.
        output_dir: Directory to extract into (will be created if needed).

    Raises:
        FileNotFoundError: If docx_path doesn't exist.
        zipfile.BadZipFile: If the file isn't a valid ZIP/DOCX.
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path, 'r') as zf:
        zf.extractall(output_dir)

    log.info("Unpacked %s → %s (%d files)", docx_path, output_dir, len(os.listdir(output_dir)))


def pack(unpacked_dir: str, output_path: str, original_docx: str | None = None) -> None:
    """Pack a directory back into a .docx file.

    Preserves the ZIP member order from the original if provided, which
    matters for some OOXML consumers (Word is picky about [Content_Types].xml
    being first).

    Args:
        unpacked_dir: Directory containing the unpacked DOCX structure.
        output_path: Path for the output .docx file.
        original_docx: Optional path to the original .docx to preserve
                       member ordering and compression settings.
    """
    if original_docx and os.path.exists(original_docx):
        _pack_preserving_order(unpacked_dir, output_path, original_docx)
    else:
        _pack_standard(unpacked_dir, output_path)

    log.info("Packed %s → %s", unpacked_dir, output_path)


def _pack_preserving_order(unpacked_dir, output_path, original_docx):
    """Pack using the member order from the original DOCX."""
    # Get ordered member list from original
    with zipfile.ZipFile(original_docx, 'r') as orig_zf:
        original_members = [info.filename for info in orig_zf.infolist()]
        original_info = {info.filename: info for info in orig_zf.infolist()}

    # Collect all files in unpacked dir
    all_files = set()
    for root, dirs, files in os.walk(unpacked_dir):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, unpacked_dir)
            # Normalize path separators to forward slash (ZIP standard)
            all_files.add(rel_path.replace(os.sep, '/'))

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        written = set()

        # First: write members in original order
        for member_name in original_members:
            if member_name in all_files:
                file_path = os.path.join(unpacked_dir, member_name.replace('/', os.sep))
                # Preserve compression type from original
                orig_info = original_info[member_name]
                zf.write(file_path, member_name, compress_type=orig_info.compress_type)
                written.add(member_name)

        # Then: write any new files not in original
        for member_name in sorted(all_files - written):
            file_path = os.path.join(unpacked_dir, member_name.replace('/', os.sep))
            zf.write(file_path, member_name)


def _pack_standard(unpacked_dir, output_path):
    """Pack with standard ordering ([Content_Types].xml first)."""
    all_files = []
    for root, dirs, files in os.walk(unpacked_dir):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, unpacked_dir).replace(os.sep, '/')
            all_files.append(rel_path)

    # Ensure [Content_Types].xml is first (OOXML requirement)
    all_files.sort(key=lambda p: (0 if p == '[Content_Types].xml' else 1, p))

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for member_name in all_files:
            file_path = os.path.join(unpacked_dir, member_name.replace('/', os.sep))
            zf.write(file_path, member_name)
