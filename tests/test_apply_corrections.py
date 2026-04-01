"""Tests for apply_corrections.py — correction application logic."""

import json
import subprocess
import sys, os
import tempfile
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from lxml import etree
from ooxml import W, IdCounter, get_runs_with_text
from apply_corrections import apply_correction, apply_all


# ── Helpers ────────────────────────────────────────────────────────────────

def _make_paragraph(texts):
    p = etree.Element(f'{W}p')
    for t in texts:
        r = etree.SubElement(p, f'{W}r')
        te = etree.SubElement(r, f'{W}t')
        te.text = t
    return p


def _make_paragraph_with_tracked_change(before, deleted, inserted, after):
    p = etree.Element(f'{W}p')
    if before:
        r = etree.SubElement(p, f'{W}r')
        t = etree.SubElement(r, f'{W}t')
        t.text = before
    d = etree.SubElement(p, f'{W}del')
    d.set(f'{W}id', '1')
    dr = etree.SubElement(d, f'{W}r')
    dt = etree.SubElement(dr, f'{W}delText')
    dt.text = deleted
    i = etree.SubElement(p, f'{W}ins')
    i.set(f'{W}id', '2')
    ir = etree.SubElement(i, f'{W}r')
    it = etree.SubElement(ir, f'{W}t')
    it.text = inserted
    if after:
        r = etree.SubElement(p, f'{W}r')
        t = etree.SubElement(r, f'{W}t')
        t.text = after
    return p


# ── apply_correction ──────────────────────────────────────────────────────

class TestApplyCorrection:
    def test_simple_replacement(self):
        p = _make_paragraph(['Ala ma kta.'])
        idc = IdCounter(100)
        assert apply_correction(p, 'kta', 'kota', 'Test', '2026-01-01T00:00:00Z', idc) is True
        assert idc.value == 102
        assert len(p.findall(f'.//{W}del')) == 1
        assert len(p.findall(f'.//{W}ins')) == 1
        assert p.findall(f'.//{W}delText')[0].text == 'kta'
        assert p.findall(f'.//{W}ins//{W}t')[0].text == 'kota'

    def test_case_insensitive_fallback(self):
        p = _make_paragraph(['Wielka Litera'])
        idc = IdCounter(100)
        assert apply_correction(p, 'wielka litera', 'mała litera', 'Test', '2026-01-01T00:00:00Z', idc) is True
        assert idc.value == 102

    def test_not_found(self):
        p = _make_paragraph(['Ala ma kota.'])
        idc = IdCounter(100)
        assert apply_correction(p, 'psa', 'kota', 'Test', '2026-01-01T00:00:00Z', idc) is False
        assert idc.value == 100

    def test_spanning_two_runs(self):
        p = _make_paragraph(['Ala m', 'a kota.'])
        idc = IdCounter(100)
        assert apply_correction(p, 'ma', 'miała', 'Test', '2026-01-01T00:00:00Z', idc) is True
        assert idc.value == 102

    def test_preserves_prefix_and_suffix(self):
        p = _make_paragraph(['Ala ma kota i psa.'])
        idc = IdCounter(100)
        apply_correction(p, 'ma', 'miała', 'Test', '2026-01-01T00:00:00Z', idc)
        all_t = [t.text for t in p.findall(f'.//{W}t') if t.text]
        all_dt = [t.text for t in p.findall(f'.//{W}delText') if t.text]
        assert 'Ala ' in all_t
        assert ' kota i psa.' in all_t
        assert 'miała' in all_t
        assert 'ma' in all_dt

    def test_empty_paragraph(self):
        p = etree.Element(f'{W}p')
        idc = IdCounter(100)
        assert apply_correction(p, 'foo', 'bar', 'Test', '2026-01-01T00:00:00Z', idc) is False

    def test_does_not_touch_existing_tracked_changes(self):
        p = _make_paragraph_with_tracked_change('Ala ', 'stare', 'nowe', ' ma kta.')
        idc = IdCounter(100)
        assert apply_correction(p, 'kta', 'kota', 'Test', '2026-01-01T00:00:00Z', idc) is True
        assert len(p.findall(f'.//{W}del')) == 2


# ── CLI + pipeline tests ─────────────────────────────────────────────────

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _make_test_docx(path: str, body_xml: str):
    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}">'
        f'<w:body>{body_xml}</w:body>'
        f'</w:document>'
    )
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
        zf.writestr('word/document.xml', doc_xml)


class TestApplyAllFunction:
    def test_apply_all_returns_counts(self, tmp_path):
        docx_path = str(tmp_path / 'input.docx')
        _make_test_docx(docx_path, '<w:p><w:r><w:t>Ala ma kta.</w:t></w:r></w:p>')

        # Unpack
        from docx_io import unpack
        unpacked = str(tmp_path / 'unpacked')
        unpack(docx_path, unpacked)

        corrections = [
            {"original": "kta", "corrected": "kota", "note": "test"},
        ]
        applied, total = apply_all(
            os.path.join(unpacked, 'word', 'document.xml'),
            corrections, 'Test', '2026-01-01T00:00:00Z',
        )
        assert applied == 1
        assert total == 1

    def test_apply_all_skips_identical(self, tmp_path):
        docx_path = str(tmp_path / 'input.docx')
        _make_test_docx(docx_path, '<w:p><w:r><w:t>Tekst.</w:t></w:r></w:p>')

        from docx_io import unpack
        unpacked = str(tmp_path / 'unpacked')
        unpack(docx_path, unpacked)

        corrections = [{"original": "Tekst", "corrected": "Tekst", "note": "no-op"}]
        applied, total = apply_all(
            os.path.join(unpacked, 'word', 'document.xml'),
            corrections, 'Test', '2026-01-01T00:00:00Z',
        )
        assert applied == 0
        assert total == 0  # identical corrections are not counted


class TestApplyCorrectionsCLI:
    def test_cli_smoke(self, tmp_path):
        docx_path = str(tmp_path / 'input.docx')
        _make_test_docx(docx_path, '<w:p><w:r><w:t>Ala ma kta.</w:t></w:r></w:p>')

        corrections_path = str(tmp_path / 'corr.json')
        with open(corrections_path, 'w') as f:
            json.dump([{"original": "kta", "corrected": "kota", "note": "fix"}], f)

        output_path = str(tmp_path / 'output.docx')
        result = subprocess.run(
            [sys.executable, '-m', 'apply_corrections', docx_path, corrections_path, '-o', output_path],
            capture_output=True, text=True,
            cwd=os.path.join(os.path.dirname(__file__), '..'),
        )
        assert result.returncode == 0
        assert os.path.exists(output_path)

        # Verify it's a valid zip
        with zipfile.ZipFile(output_path) as zf:
            assert 'word/document.xml' in zf.namelist()


class TestFullPipeline:
    def test_extract_apply_verify(self, tmp_path):
        """Full pipeline: extract → (mock corrections) → apply → verify."""
        docx_path = str(tmp_path / 'input.docx')
        _make_test_docx(docx_path, '<w:p><w:r><w:t>Ala ma kta i psa.</w:t></w:r></w:p>')

        corrections = [{"original": "kta", "corrected": "kota", "note": "fix"}]
        corrections_path = str(tmp_path / 'corr.json')
        with open(corrections_path, 'w') as f:
            json.dump(corrections, f)

        output_path = str(tmp_path / 'output.docx')
        result = subprocess.run(
            [sys.executable, '-m', 'apply_corrections', docx_path, corrections_path, '-o', output_path],
            capture_output=True, text=True,
            cwd=os.path.join(os.path.dirname(__file__), '..'),
        )
        assert result.returncode == 0

        # Verify: original(visible) vs corrected(original) should match
        from verify_docx import compare
        exit_code = compare(docx_path, output_path, quiet=True)
        assert exit_code == 0
