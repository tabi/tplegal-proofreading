"""Tests for extract_text.py — text extraction from DOCX."""

import os
import sys
import subprocess
import tempfile
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from extract_text import extract_text


# ── Helpers ────────────────────────────────────────────────────────────────

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _make_docx(tmp_path, body_xml: str) -> str:
    """Create a minimal .docx with the given body XML content."""
    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}">'
        f'<w:body>{body_xml}</w:body>'
        f'</w:document>'
    )
    path = str(tmp_path / 'test.docx')
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
        zf.writestr('word/document.xml', doc_xml)
    return path


# ── Tests ──────────────────────────────────────────────────────────────────

class TestExtractText:
    def test_simple_paragraph(self, tmp_path):
        docx = _make_docx(tmp_path, '<w:p><w:r><w:t>Ala ma kota.</w:t></w:r></w:p>')
        lines = extract_text(docx)
        assert len(lines) == 1
        assert lines[0] == '\u00b6001: Ala ma kota.'

    def test_multiple_paragraphs_numbered(self, tmp_path):
        body = (
            '<w:p><w:r><w:t>Pierwszy.</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Drugi.</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Trzeci.</w:t></w:r></w:p>'
        )
        docx = _make_docx(tmp_path, body)
        lines = extract_text(docx)
        assert len(lines) == 3
        assert lines[0].startswith('\u00b6001:')
        assert lines[1].startswith('\u00b6002:')
        assert lines[2].startswith('\u00b6003:')

    def test_skips_empty_paragraphs(self, tmp_path):
        body = (
            '<w:p><w:r><w:t>Tekst.</w:t></w:r></w:p>'
            '<w:p/>'
            '<w:p><w:r><w:t> </w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Drugi.</w:t></w:r></w:p>'
        )
        docx = _make_docx(tmp_path, body)
        lines = extract_text(docx)
        assert len(lines) == 2

    def test_skips_field_codes(self, tmp_path):
        body = (
            '<w:p>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            '</w:p>'
            '<w:p><w:r><w:t>Tekst.</w:t></w:r></w:p>'
        )
        docx = _make_docx(tmp_path, body)
        lines = extract_text(docx)
        assert len(lines) == 1
        assert 'Tekst.' in lines[0]

    def test_skips_tracked_changes(self, tmp_path):
        body = (
            '<w:p>'
            '<w:r><w:t>Przed </w:t></w:r>'
            '<w:del w:id="1" w:author="X" w:date="2026-01-01T00:00:00Z">'
            '<w:r><w:delText>usuniety</w:delText></w:r>'
            '</w:del>'
            '<w:ins w:id="2" w:author="X" w:date="2026-01-01T00:00:00Z">'
            '<w:r><w:t>wstawiony</w:t></w:r>'
            '</w:ins>'
            '<w:r><w:t> po.</w:t></w:r>'
            '</w:p>'
        )
        docx = _make_docx(tmp_path, body)
        lines = extract_text(docx)
        assert len(lines) == 1
        # visible mode: skip del, include ins
        assert 'usuniety' not in lines[0]
        assert 'wstawiony' in lines[0]

    def test_table_paragraphs(self, tmp_path):
        body = (
            '<w:tbl><w:tr><w:tc>'
            '<w:p><w:r><w:t>Komorka.</w:t></w:r></w:p>'
            '</w:tc></w:tr></w:tbl>'
        )
        docx = _make_docx(tmp_path, body)
        lines = extract_text(docx)
        assert len(lines) == 1
        assert 'Komorka.' in lines[0]

    def test_file_not_found(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            extract_text(str(tmp_path / 'nonexistent.docx'))

    def test_cli_smoke(self, tmp_path):
        docx = _make_docx(tmp_path, '<w:p><w:r><w:t>CLI test.</w:t></w:r></w:p>')
        result = subprocess.run(
            [sys.executable, '-m', 'extract_text', docx],
            capture_output=True, text=True,
            cwd=os.path.join(os.path.dirname(__file__), '..'),
        )
        assert result.returncode == 0
        assert 'CLI test.' in result.stdout
