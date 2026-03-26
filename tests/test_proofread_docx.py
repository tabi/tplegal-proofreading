"""Tests for proofread_docx.py — multi-correction logic and integration."""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from lxml import etree

from ooxml import W, IdCounter
from proofread_docx import (
    _apply_single_match,
    apply_corrections_to_paragraph,
)


# ── Helpers ────────────────────────────────────────────────────────────────

def _make_paragraph(texts):
    p = etree.Element(f'{W}p')
    for t in texts:
        r = etree.SubElement(p, f'{W}r')
        te = etree.SubElement(r, f'{W}t')
        te.text = t
    return p


# ── _apply_single_match ──────────────────────────────────────────────────

class TestApplySingleMatch:
    def test_applies(self):
        p = _make_paragraph(['Ala ma kta.'])
        c = IdCounter(100)
        match = {'offset': 7, 'length': 3, 'end': 10, 'replacement': 'kota',
                 'message': '', 'rule_id': '', 'category': ''}
        assert _apply_single_match(p, match, c, 'T', '2026-01-01T00:00:00Z') is True
        assert len(p.findall(f'.//{W}del')) == 1

    def test_out_of_range(self):
        p = _make_paragraph(['short'])
        c = IdCounter(100)
        match = {'offset': 0, 'length': 100, 'end': 100, 'replacement': 'x',
                 'message': '', 'rule_id': '', 'category': ''}
        assert _apply_single_match(p, match, c, 'T', 'D') is False

    def test_empty_paragraph(self):
        p = etree.Element(f'{W}p')
        c = IdCounter(100)
        match = {'offset': 0, 'length': 1, 'end': 1, 'replacement': 'x',
                 'message': '', 'rule_id': '', 'category': ''}
        assert _apply_single_match(p, match, c, 'T', 'D') is False


# ── apply_corrections_to_paragraph ────────────────────────────────────────

class TestApplyCorrections:
    def test_single(self):
        p = _make_paragraph(['Ala ma kta.'])
        c = IdCounter(100)
        matches = [{'offset': 7, 'length': 3, 'end': 10, 'replacement': 'kota',
                     'message': '', 'rule_id': '', 'category': ''}]
        assert apply_corrections_to_paragraph(p, matches, c, 'T', 'D') == 1

    def test_multiple_in_separate_runs(self):
        p = _make_paragraph(['Aa', ' ma ', 'kta', '.'])
        c = IdCounter(100)
        matches = [
            {'offset': 0, 'length': 2, 'end': 2, 'replacement': 'Ala',
             'message': '', 'rule_id': '', 'category': ''},
            {'offset': 6, 'length': 3, 'end': 9, 'replacement': 'kota',
             'message': '', 'rule_id': '', 'category': ''},
        ]
        assert apply_corrections_to_paragraph(p, matches, c, 'T', 'D') == 2

    def test_multiple_in_single_run(self):
        """Bug fix: two corrections in the same run should both apply."""
        p = _make_paragraph(['Aa ma kta.'])
        c = IdCounter(100)
        matches = [
            {'offset': 0, 'length': 2, 'end': 2, 'replacement': 'Ala',
             'message': '', 'rule_id': '', 'category': ''},
            {'offset': 6, 'length': 3, 'end': 9, 'replacement': 'kota',
             'message': '', 'rule_id': '', 'category': ''},
        ]
        n = apply_corrections_to_paragraph(p, matches, c, 'T', 'D')
        assert n == 2
        assert len(p.findall(f'.//{W}del')) == 2
        assert len(p.findall(f'.//{W}ins')) == 2

    def test_no_matches(self):
        p = _make_paragraph(['Clean.'])
        c = IdCounter(100)
        assert apply_corrections_to_paragraph(p, [], c, 'T', 'D') == 0

    def test_empty_paragraph(self):
        p = etree.Element(f'{W}p')
        c = IdCounter(100)
        match = {'offset': 0, 'length': 3, 'end': 3, 'replacement': 'x',
                 'message': '', 'rule_id': '', 'category': ''}
        assert apply_corrections_to_paragraph(p, [match], c, 'T', 'D') == 0
