"""Tests for docx_io.py — DOCX pack/unpack."""

import os
import sys
import tempfile
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx_io import unpack, pack


def _create_fake_docx(path):
    """Create a minimal valid .docx (ZIP with [Content_Types].xml and word/document.xml)."""
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
        zf.writestr('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>')
        zf.writestr('word/_rels/document.xml.rels', '<?xml version="1.0"?><Relationships/>')


class TestUnpack:
    def test_unpacks_docx(self, tmp_path):
        docx = tmp_path / 'test.docx'
        _create_fake_docx(str(docx))

        out_dir = tmp_path / 'unpacked'
        unpack(str(docx), str(out_dir))

        assert (out_dir / '[Content_Types].xml').exists()
        assert (out_dir / 'word' / 'document.xml').exists()

    def test_missing_file(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            unpack(str(tmp_path / 'nonexistent.docx'), str(tmp_path / 'out'))

    def test_invalid_zip(self, tmp_path):
        bad_file = tmp_path / 'bad.docx'
        bad_file.write_text('not a zip')
        with pytest.raises(zipfile.BadZipFile):
            unpack(str(bad_file), str(tmp_path / 'out'))


class TestPack:
    def test_roundtrip(self, tmp_path):
        """Unpack then repack should produce a valid ZIP."""
        original = tmp_path / 'original.docx'
        _create_fake_docx(str(original))

        unpacked = tmp_path / 'unpacked'
        unpack(str(original), str(unpacked))

        repacked = tmp_path / 'repacked.docx'
        pack(str(unpacked), str(repacked), original_docx=str(original))

        assert repacked.exists()
        with zipfile.ZipFile(str(repacked), 'r') as zf:
            names = zf.namelist()
            assert '[Content_Types].xml' in names
            assert 'word/document.xml' in names

    def test_preserves_member_order(self, tmp_path):
        original = tmp_path / 'original.docx'
        _create_fake_docx(str(original))

        with zipfile.ZipFile(str(original), 'r') as zf:
            original_order = zf.namelist()

        unpacked = tmp_path / 'unpacked'
        unpack(str(original), str(unpacked))

        repacked = tmp_path / 'repacked.docx'
        pack(str(unpacked), str(repacked), original_docx=str(original))

        with zipfile.ZipFile(str(repacked), 'r') as zf:
            repacked_order = zf.namelist()

        assert repacked_order == original_order

    def test_pack_without_original(self, tmp_path):
        """Pack without original should still produce valid DOCX."""
        unpacked = tmp_path / 'unpacked' / 'word'
        unpacked.mkdir(parents=True)
        (tmp_path / 'unpacked' / '[Content_Types].xml').write_text('<Types/>')
        (unpacked / 'document.xml').write_text('<doc/>')

        output = tmp_path / 'output.docx'
        pack(str(tmp_path / 'unpacked'), str(output))

        with zipfile.ZipFile(str(output), 'r') as zf:
            names = zf.namelist()
            # [Content_Types].xml should be first
            assert names[0] == '[Content_Types].xml'

    def test_new_files_added(self, tmp_path):
        """Files added after unpacking should be included."""
        original = tmp_path / 'original.docx'
        _create_fake_docx(str(original))

        unpacked = tmp_path / 'unpacked'
        unpack(str(original), str(unpacked))

        # Add a new file
        (unpacked / 'word' / 'comments.xml').write_text('<comments/>')

        repacked = tmp_path / 'repacked.docx'
        pack(str(unpacked), str(repacked), original_docx=str(original))

        with zipfile.ZipFile(str(repacked), 'r') as zf:
            assert 'word/comments.xml' in zf.namelist()
