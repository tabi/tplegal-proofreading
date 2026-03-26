"""Shared OOXML helpers for tracked-change proofreading tools."""

from __future__ import annotations

import copy
import logging
from lxml import etree

log = logging.getLogger(__name__)

# ── Namespace ──────────────────────────────────────────────────────────────

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


# ── Tracked-change ID counter ─────────────────────────────────────────────

class IdCounter:
    """Sequential ID generator for tracked-change elements."""

    def __init__(self, start: int = 0) -> None:
        self._next = start

    def next(self) -> int:
        self._next += 1
        return self._next

    @property
    def value(self):
        return self._next

    @value.setter
    def value(self, v):
        self._next = v


def find_max_id(tree: etree._ElementTree) -> int:
    """Find the maximum w:id already present in an ElementTree."""
    max_id = 0
    for elem in tree.iter():
        val = elem.get(f'{W}id')
        if val is not None:
            try:
                max_id = max(max_id, int(val))
            except ValueError:
                pass
    return max_id


# ── Run helpers ────────────────────────────────────────────────────────────

def make_run(text: str, rpr: etree._Element | None = None, is_delete: bool = False) -> etree._Element:
    """Create a w:r element with optional formatting (w:rPr deep-copied)."""
    r = etree.Element(f'{W}r')
    if rpr is not None:
        r.append(copy.deepcopy(rpr))

    tag = f'{W}delText' if is_delete else f'{W}t'
    t = etree.SubElement(r, tag)
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set(XML_SPACE, 'preserve')

    return r


def get_rpr(run: etree._Element) -> etree._Element | None:
    """Extract w:rPr element from a run, or None."""
    return run.find(f'{W}rPr')


# ── Tracked-change element builders ───────────────────────────────────────

def make_del(text: str, rpr: etree._Element | None, *, id_counter: IdCounter, author: str, date_str: str) -> etree._Element:
    """Create a w:del tracked-change element."""
    el = etree.Element(f'{W}del')
    el.set(f'{W}id', str(id_counter.next()))
    el.set(f'{W}author', author)
    el.set(f'{W}date', date_str)
    el.append(make_run(text, rpr, is_delete=True))
    return el


def make_ins(text: str, rpr: etree._Element | None, *, id_counter: IdCounter, author: str, date_str: str) -> etree._Element:
    """Create a w:ins tracked-change element."""
    el = etree.Element(f'{W}ins')
    el.set(f'{W}id', str(id_counter.next()))
    el.set(f'{W}author', author)
    el.set(f'{W}date', date_str)
    el.append(make_run(text, rpr))
    return el


# ── Paragraph text extraction ─────────────────────────────────────────────

def _is_inside_tracked_change(elem):
    """Check if element is nested inside w:del or w:ins at any depth."""
    parent = elem.getparent()
    while parent is not None:
        if parent.tag in (f'{W}del', f'{W}ins'):
            return True
        parent = parent.getparent()
    return False


def get_paragraph_runs(p_elem: etree._Element) -> list[dict]:
    """Extract text runs from a paragraph as list of dicts.

    Only includes direct w:r children; skips runs inside w:ins, w:del,
    and special runs (field codes, comment refs).

    Each dict has keys: elem, text, start, end, rpr.
    """
    runs = []
    offset = 0
    for child in p_elem:
        if child.tag != f'{W}r':
            continue
        # Skip special runs
        if child.find(f'{W}fldChar') is not None:
            continue
        if child.find(f'{W}instrText') is not None:
            continue
        if child.find(f'{W}commentReference') is not None:
            continue

        text_parts = []
        for t_elem in child.findall(f'{W}t'):
            if t_elem.text:
                text_parts.append(t_elem.text)
        text = ''.join(text_parts)

        rpr = child.find(f'{W}rPr')
        runs.append({
            'elem': child,
            'text': text,
            'start': offset,
            'end': offset + len(text),
            'rpr': copy.deepcopy(rpr) if rpr is not None else None,
        })
        offset += len(text)
    return runs


def get_runs_with_text(para: etree._Element) -> list[tuple[etree._Element, str, int, int]]:
    """Get list of (run_element, text, start_pos, end_pos) for text-bearing runs.

    Wrapper around get_paragraph_runs() for backwards compatibility.
    Returns tuples instead of dicts.
    """
    return [
        (r['elem'], r['text'], r['start'], r['end'])
        for r in get_paragraph_runs(para)
    ]


def get_paragraph_text_from_runs(runs: list[dict]) -> str:
    """Get plain text from a run-dict list (as returned by get_paragraph_runs)."""
    return ''.join(r['text'] for r in runs)


def get_paragraph_text(para: etree._Element) -> str:
    """Extract concatenated text from active runs."""
    return get_paragraph_text_from_runs(get_paragraph_runs(para))
