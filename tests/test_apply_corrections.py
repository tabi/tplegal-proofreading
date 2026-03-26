"""Tests for apply_corrections.py — correction application logic."""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from lxml import etree
from ooxml import W, IdCounter, get_runs_with_text
from apply_corrections import apply_correction


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
