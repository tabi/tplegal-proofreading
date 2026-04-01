"""Tests for verify_docx.py and minidom_helpers mode parameter."""

import os
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import defusedxml.minidom as minidom_mod
from minidom_helpers import extract_paragraph_text, find_elements
from verify_docx import extract_stats, compare


# ── Helpers ────────────────────────────────────────────────────────────────

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _make_docx(tmp_path, body_xml: str, name: str = 'test.docx') -> str:
    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}">'
        f'<w:body>{body_xml}</w:body>'
        f'</w:document>'
    )
    path = str(tmp_path / name)
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
        zf.writestr('word/document.xml', doc_xml)
    return path


def _parse_para(xml_str: str):
    """Parse a paragraph XML string and return the first p element."""
    full = f'<w:document xmlns:w="{W_NS}"><w:body>{xml_str}</w:body></w:document>'
    dom = minidom_mod.parseString(full)
    return find_elements(dom.documentElement, "p")[0]


# ── minidom_helpers mode tests ─────────────────────────────────────────────

class TestExtractParagraphTextModes:
    def test_visible_skips_del_text(self):
        p = _parse_para(
            '<w:p>'
            '<w:r><w:t>Przed </w:t></w:r>'
            '<w:del w:id="1" w:author="X" w:date="D">'
            '<w:r><w:delText>usuniety</w:delText></w:r>'
            '</w:del>'
            '<w:ins w:id="2" w:author="X" w:date="D">'
            '<w:r><w:t>wstawiony</w:t></w:r>'
            '</w:ins>'
            '<w:r><w:t> po.</w:t></w:r>'
            '</w:p>'
        )
        text = extract_paragraph_text(p, mode='visible')
        assert text == 'Przed wstawiony po.'
        assert 'usuniety' not in text

    def test_original_skips_ins_text(self):
        p = _parse_para(
            '<w:p>'
            '<w:r><w:t>Przed </w:t></w:r>'
            '<w:del w:id="1" w:author="X" w:date="D">'
            '<w:r><w:delText>usuniety</w:delText></w:r>'
            '</w:del>'
            '<w:ins w:id="2" w:author="X" w:date="D">'
            '<w:r><w:t>wstawiony</w:t></w:r>'
            '</w:ins>'
            '<w:r><w:t> po.</w:t></w:r>'
            '</w:p>'
        )
        text = extract_paragraph_text(p, mode='original')
        assert text == 'Przed usuniety po.'
        assert 'wstawiony' not in text

    def test_default_mode_is_visible(self):
        p = _parse_para(
            '<w:p>'
            '<w:del w:id="1" w:author="X" w:date="D">'
            '<w:r><w:delText>gone</w:delText></w:r>'
            '</w:del>'
            '<w:r><w:t>stays</w:t></w:r>'
            '</w:p>'
        )
        assert extract_paragraph_text(p) == 'stays'

    def test_no_tracked_changes_both_modes_equal(self):
        p = _parse_para('<w:p><w:r><w:t>Prosty tekst.</w:t></w:r></w:p>')
        assert extract_paragraph_text(p, 'visible') == extract_paragraph_text(p, 'original')
        assert extract_paragraph_text(p, 'visible') == 'Prosty tekst.'


# ── verify_docx false positive fix ────────────────────────────────────────

class TestVerifyDocxTrackedChanges:
    def test_no_false_positive_after_correction(self, tmp_path):
        """Original vs. corrected with tracked changes should PASS."""
        orig_body = '<w:p><w:r><w:t>Ala ma kta.</w:t></w:r></w:p>'
        corr_body = (
            '<w:p>'
            '<w:r><w:t>Ala ma </w:t></w:r>'
            '<w:del w:id="1" w:author="K" w:date="D">'
            '<w:r><w:delText>kta</w:delText></w:r>'
            '</w:del>'
            '<w:ins w:id="2" w:author="K" w:date="D">'
            '<w:r><w:t>kota</w:t></w:r>'
            '</w:ins>'
            '<w:r><w:t>.</w:t></w:r>'
            '</w:p>'
        )
        orig_path = _make_docx(tmp_path, orig_body, 'orig.docx')
        corr_path = _make_docx(tmp_path, corr_body, 'corr.docx')

        # orig visible = "Ala ma kta.", corr original = "Ala ma kta."
        exit_code = compare(orig_path, corr_path, quiet=True)
        assert exit_code == 0

    def test_true_positive_on_damaged_doc(self, tmp_path):
        """Original vs. genuinely damaged doc should FAIL."""
        orig_body = (
            '<w:p><w:r><w:t>Pierwszy akapit.</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Drugi akapit.</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Trzeci akapit z dlugim tekstem do porownania.</w:t></w:r></w:p>'
        )
        damaged_body = '<w:p><w:r><w:t>Pierwszy akapit.</w:t></w:r></w:p>'
        orig_path = _make_docx(tmp_path, orig_body, 'orig.docx')
        dmg_path = _make_docx(tmp_path, damaged_body, 'dmg.docx')

        exit_code = compare(orig_path, dmg_path, quiet=True)
        assert exit_code >= 1  # WARNING or ERROR

    def test_extract_stats_mode_visible(self, tmp_path):
        body = (
            '<w:p>'
            '<w:r><w:t>Widoczny</w:t></w:r>'
            '<w:del w:id="1" w:author="X" w:date="D">'
            '<w:r><w:delText>usuniety</w:delText></w:r>'
            '</w:del>'
            '</w:p>'
        )
        path = _make_docx(tmp_path, body)
        stats = extract_stats(path, mode='visible')
        assert stats.paragraphs[0] == 'Widoczny'

    def test_extract_stats_mode_original(self, tmp_path):
        body = (
            '<w:p>'
            '<w:r><w:t>Widoczny</w:t></w:r>'
            '<w:ins w:id="1" w:author="X" w:date="D">'
            '<w:r><w:t>dodany</w:t></w:r>'
            '</w:ins>'
            '</w:p>'
        )
        path = _make_docx(tmp_path, body)
        stats = extract_stats(path, mode='original')
        assert stats.paragraphs[0] == 'Widoczny'
        assert 'dodany' not in stats.paragraphs[0]
