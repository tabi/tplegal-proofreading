"""Tests for ooxml.py — shared OOXML helpers."""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from lxml import etree
from ooxml import (
    W,
    IdCounter,
    find_max_id,
    make_run,
    make_del,
    make_ins,
    get_rpr,
    _is_inside_tracked_change,
    get_paragraph_runs,
    get_runs_with_text,
    get_paragraph_text,
    get_paragraph_text_from_runs,
)


# ── Helpers ────────────────────────────────────────────────────────────────

def _make_paragraph(texts, bold_indices=None):
    bold_indices = bold_indices or set()
    p = etree.Element(f'{W}p')
    for i, t in enumerate(texts):
        r = etree.SubElement(p, f'{W}r')
        if i in bold_indices:
            rpr = etree.SubElement(r, f'{W}rPr')
            etree.SubElement(rpr, f'{W}b')
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


# ── IdCounter ─────────────────────────────────────────────────────────────

class TestIdCounter:
    def test_sequential(self):
        c = IdCounter(10)
        assert c.next() == 11
        assert c.next() == 12

    def test_value_property(self):
        c = IdCounter(5)
        assert c.value == 5
        c.next()
        assert c.value == 6

    def test_value_setter(self):
        c = IdCounter(0)
        c.value = 100
        assert c.next() == 101


# ── find_max_id ───────────────────────────────────────────────────────────

class TestFindMaxId:
    def test_finds_max(self):
        root = etree.Element(f'{W}body')
        d = etree.SubElement(root, f'{W}del')
        d.set(f'{W}id', '42')
        i = etree.SubElement(root, f'{W}ins')
        i.set(f'{W}id', '99')
        assert find_max_id(etree.ElementTree(root)) == 99

    def test_empty(self):
        root = etree.Element(f'{W}body')
        assert find_max_id(etree.ElementTree(root)) == 0

    def test_non_numeric(self):
        root = etree.Element(f'{W}body')
        p = etree.SubElement(root, f'{W}p')
        p.set(f'{W}id', 'abc')
        assert find_max_id(etree.ElementTree(root)) == 0


# ── make_run ──────────────────────────────────────────────────────────────

class TestMakeRun:
    def test_regular(self):
        r = make_run('test')
        assert r.tag == f'{W}r'
        assert r.find(f'{W}t').text == 'test'
        assert r.getparent() is None

    def test_delete(self):
        r = make_run('gone', is_delete=True)
        assert r.find(f'{W}delText').text == 'gone'
        assert r.find(f'{W}t') is None

    def test_space_preserve(self):
        r = make_run(' spaced ')
        t = r.find(f'{W}t')
        assert t.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve'

    def test_no_space_preserve(self):
        r = make_run('normal')
        assert r.find(f'{W}t').get('{http://www.w3.org/XML/1998/namespace}space') is None

    def test_copies_formatting(self):
        rpr = etree.Element(f'{W}rPr')
        etree.SubElement(rpr, f'{W}b')
        r = make_run('bold', rpr=rpr)
        result_rpr = r.find(f'{W}rPr')
        assert result_rpr is not None
        assert result_rpr.find(f'{W}b') is not None


# ── make_del / make_ins ───────────────────────────────────────────────────

class TestMakeDelIns:
    def test_del(self):
        c = IdCounter(100)
        d = make_del('removed', None, id_counter=c, author='A', date_str='2026-01-01T00:00:00Z')
        assert d.tag == f'{W}del'
        assert d.get(f'{W}author') == 'A'
        assert d.find(f'.//{W}delText').text == 'removed'

    def test_ins(self):
        c = IdCounter(100)
        i = make_ins('added', None, id_counter=c, author='A', date_str='2026-01-01T00:00:00Z')
        assert i.tag == f'{W}ins'
        assert i.find(f'.//{W}t').text == 'added'

    def test_ids_increment(self):
        c = IdCounter(100)
        d1 = make_del('a', None, id_counter=c, author='A', date_str='D')
        d2 = make_del('b', None, id_counter=c, author='A', date_str='D')
        assert int(d2.get(f'{W}id')) > int(d1.get(f'{W}id'))


# ── _is_inside_tracked_change ─────────────────────────────────────────────

class TestIsInsideTrackedChange:
    def test_direct_child_of_del(self):
        p = etree.Element(f'{W}p')
        d = etree.SubElement(p, f'{W}del')
        r = etree.SubElement(d, f'{W}r')
        assert _is_inside_tracked_change(r) is True

    def test_nested_deeper(self):
        p = etree.Element(f'{W}p')
        ins = etree.SubElement(p, f'{W}ins')
        hl = etree.SubElement(ins, f'{W}hyperlink')
        r = etree.SubElement(hl, f'{W}r')
        assert _is_inside_tracked_change(r) is True

    def test_normal_run(self):
        p = etree.Element(f'{W}p')
        r = etree.SubElement(p, f'{W}r')
        assert _is_inside_tracked_change(r) is False


# ── get_paragraph_runs ────────────────────────────────────────────────────

class TestGetParagraphRuns:
    def test_simple(self):
        p = _make_paragraph(['Hello ', 'world'])
        runs = get_paragraph_runs(p)
        assert len(runs) == 2
        assert runs[0]['text'] == 'Hello '
        assert runs[0]['start'] == 0
        assert runs[0]['end'] == 6
        assert runs[1]['text'] == 'world'

    def test_skips_field_chars(self):
        p = _make_paragraph(['text'])
        special_r = etree.SubElement(p, f'{W}r')
        etree.SubElement(special_r, f'{W}fldChar')
        assert len(get_paragraph_runs(p)) == 1

    def test_skips_tracked_changes(self):
        p = etree.Element(f'{W}p')
        r1 = etree.SubElement(p, f'{W}r')
        t1 = etree.SubElement(r1, f'{W}t')
        t1.text = 'visible'
        d = etree.SubElement(p, f'{W}del')
        r2 = etree.SubElement(d, f'{W}r')
        t2 = etree.SubElement(r2, f'{W}t')
        t2.text = 'deleted'
        runs = get_paragraph_runs(p)
        assert len(runs) == 1
        assert runs[0]['text'] == 'visible'

    def test_preserves_formatting(self):
        p = _make_paragraph(['normal', 'bold'], bold_indices={1})
        runs = get_paragraph_runs(p)
        assert runs[0]['rpr'] is None
        assert runs[1]['rpr'] is not None

    def test_empty(self):
        assert get_paragraph_runs(etree.Element(f'{W}p')) == []


# ── get_runs_with_text ────────────────────────────────────────────────────

class TestGetRunsWithText:
    def test_simple(self):
        p = _make_paragraph(['Ala ', 'ma ', 'kota.'])
        runs = get_runs_with_text(p)
        assert [r[1] for r in runs] == ['Ala ', 'ma ', 'kota.']

    def test_skips_tracked_changes(self):
        p = _make_paragraph_with_tracked_change('Ala ', 'ma', 'miała', 'kota.')
        runs = get_runs_with_text(p)
        assert [r[1] for r in runs] == ['Ala ', 'kota.']

    def test_empty(self):
        assert get_runs_with_text(etree.Element(f'{W}p')) == []


# ── get_paragraph_text ────────────────────────────────────────────────────

class TestGetParagraphText:
    def test_concatenation(self):
        p = _make_paragraph(['Ala ', 'ma ', 'kota.'])
        assert get_paragraph_text(p) == 'Ala ma kota.'

    def test_ignores_tracked_changes(self):
        p = _make_paragraph_with_tracked_change('Przed ', 'usunięte', 'wstawione', ' po')
        assert get_paragraph_text(p) == 'Przed  po'

    def test_from_runs(self):
        runs = [{'text': 'a'}, {'text': 'b'}]
        assert get_paragraph_text_from_runs(runs) == 'ab'
