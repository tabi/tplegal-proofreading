#!/usr/bin/env python3
"""
Nanoszenie korekt z corrections.json na .docx jako śledzone zmiany (w:del/w:ins).

Gwarancje v2.3 (audyt 24.09.2026):
  - dopasowanie DOKŁADNE (bez fallbacku na wielkość liter) na tym samym tekście,
    który pokazuje extract-text (wspólny docmodel),
  - `original` występujący więcej niż raz → korekta ODRZUCONA, chyba że pole
    `paragraph` (numer ¶ z extract-text) wskazuje jedno wystąpienie,
  - śledzona zmiana obejmuje tylko wyrazy/znaki faktycznie różne,
  - tabulatory, przypisy, obrazy, pola, cudze śledzone zmiany i formatowanie
    są nietykalne: korekta, która by je naruszyła, jest ODRZUCANA z powodem,
  - korekta zmieniająca cyfry albo dotykająca „&”, „@”, „/”, „\\”, „_” jest
    ODRZUCANA (sygnatury, numery, daty, kwoty, nazwy, adresy),
  - raport na stdout wymienia KAŻDĄ korektę ze statusem.

Format corrections.json:
[
  {"original": "od tego czy", "corrected": "od tego, czy", "note": "interpunkcja", "paragraph": 12}
]
`paragraph` jest opcjonalne (liczba albo "¶012").

Usage:
  apply-corrections input.docx corrections.json -o output.docx [--author "Korektor AI"] [--report r.json]

Kody wyjścia: 0 = wszystkie naniesione, 1 = część odrzucona/nieznaleziona, 2 = błąd / nic nie naniesiono.
"""

from __future__ import annotations

import argparse
import copy
import difflib
import glob
import json
import logging
import os
import re
import sys
import tempfile
from datetime import datetime, timezone

from lxml import etree

from docmodel import (
    W, XML_SPACE, Document, Paragraph, max_annotation_id, parse_xml, rpr_signature, scan_paragraph,
)
from ooxml import IdCounter, make_ins

log = logging.getLogger(__name__)

TOKEN_RE = re.compile(r'\w+|\s+|[^\w\s]')
PROTECTED_CHARS = set('&@/\\_')

STATUS_APPLIED = 'applied'
STATUS_REJECTED = 'rejected'
STATUS_NOT_FOUND = 'not_found'
STATUS_NOOP = 'noop'


class CorrectionsError(ValueError):
    """Niepoprawny plik corrections.json."""


# ── Walidacja wejścia ─────────────────────────────────────────────────────

def _parse_paragraph_ref(value) -> int | None:
    if value is None or value == '':
        return None
    if isinstance(value, bool):
        raise CorrectionsError(f'pole "paragraph" ma zły typ: {value!r}')
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        m = re.fullmatch(r'\s*¶?\s*(\d+)\s*', value)
        if m:
            return int(m.group(1))
    raise CorrectionsError(f'pole "paragraph" musi być numerem akapitu (np. 12 albo "¶012"), jest: {value!r}')


def validate_corrections(data) -> list[dict]:
    if not isinstance(data, list):
        raise CorrectionsError('corrections.json musi być tablicą JSON')
    out = []
    for i, corr in enumerate(data, 1):
        if not isinstance(corr, dict):
            raise CorrectionsError(f'korekta #{i}: musi być obiektem JSON')
        original, corrected = corr.get('original'), corr.get('corrected')
        if not isinstance(original, str) or not original:
            raise CorrectionsError(f'korekta #{i}: brak niepustego pola "original"')
        if not isinstance(corrected, str):
            raise CorrectionsError(f'korekta #{i}: brak pola "corrected"')
        out.append({
            'original': original,
            'corrected': corrected,
            'note': str(corr.get('note') or ''),
            'paragraph': _parse_paragraph_ref(corr.get('paragraph')),
        })
    return out


# ── Plan zmian: tylko różniące się tokeny ─────────────────────────────────

def plan_edits(original: str, corrected: str) -> list[tuple[int, int, str]]:
    """Lista (start, end, wstawka) względem `original` — minimalne różnice na poziomie wyrazów."""
    a = TOKEN_RE.findall(original)
    b = TOKEN_RE.findall(corrected)
    pos = [0]
    for tok in a:
        pos.append(pos[-1] + len(tok))
    edits = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag != 'equal':
            edits.append((pos[i1], pos[i2], ''.join(b[j1:j2])))
    return edits


def _find_occurrences(doc: Document, needle: str) -> list[tuple[Paragraph, int]]:
    hits = []
    for para in doc.numbered():
        text = para.text
        start = text.find(needle)
        while start != -1:
            hits.append((para, start))
            start = text.find(needle, start + 1)
    return hits


# ── Walidacja jednej zmiany ───────────────────────────────────────────────

def _check_content(text: str, s: int, e: int, ins: str) -> str | None:
    deleted = text[s:e]
    if re.findall(r'\d', deleted) != re.findall(r'\d', ins):
        return 'korekta zmienia cyfry (numery, daty, kwoty, sygnatury) — zakazane'
    neighbours = (text[s - 1] if s > 0 else '') + (text[e] if e < len(text) else '')
    touched = (set(deleted) | set(ins) | set(neighbours)) & PROTECTED_CHARS
    if touched:
        chars = ', '.join(f'„{c}”' for c in sorted(touched))
        return f'korekta przy znaku {chars} (nazwy, adresy, sygnatury) — zakazane'
    return None


def _insertion_anchor(para: Paragraph, s: int):
    """Run, do którego dokleić wstawkę w pozycji s: (item, 'after'|'before') albo (None, powód)."""
    before = [i for i in para.items if i.kind == 'text' and i.start < s <= i.end]
    after = [i for i in para.items if i.kind == 'text' and i.start == s and i.end > s]
    for item, side in [(x, 'after') for x in before] + [(x, 'before') for x in after]:
        if item.lock is None:
            return item, side
    locked = [i.lock for i in before + after if i.lock]
    if locked:
        return None, locked[0]
    return None, 'brak tekstu, do którego można dokleić wstawkę (sąsiaduje z elementem nietekstowym)'


def _check_span(para: Paragraph, s: int, e: int, ins: str) -> str | None:
    if e == s:
        anchor, reason = _insertion_anchor(para, s)
        return None if anchor is not None else reason
    overlapping = [i for i in para.items if i.start < e and i.end > s]
    inside_zero = [i for i in para.items if i.start == i.end and s < i.start < e]
    blockers = [i for i in overlapping + inside_zero if i.kind == 'token']
    if blockers:
        names = ', '.join(sorted({b.name or 'element' for b in blockers}))
        return f'fragment zawiera element nietekstowy ({names}) — zawęź "original" tak, by go nie obejmował'
    locks = [i.lock for i in overlapping if i.lock]
    if locks:
        return locks[0]
    if len({id(i.container) for i in overlapping}) > 1:
        return 'fragment przechodzi przez granicę hiperłącza/kontrolki — rozbij korektę'
    if len({rpr_signature(i.run.find(f'{W}rPr')) for i in overlapping}) > 1:
        return 'fragment ma mieszane formatowanie (np. część pogrubiona) — rozbij korektę na części o jednym formatowaniu'
    return None


# ── Operacje na drzewie ───────────────────────────────────────────────────

def _content_children(run) -> list:
    return [c for c in run if isinstance(c.tag, str) and c.tag != f'{W}rPr']


def _split_run_after(run, child):
    """Przenieś wszystko po `child` do nowego runu tuż za `run` (z kopią w:rPr)."""
    following = []
    node = child.getnext()
    while node is not None:
        following.append(node)
        node = node.getnext()
    if not following:
        return None
    new_run = etree.Element(run.tag, attrib=dict(run.attrib))
    rpr = run.find(f'{W}rPr')
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    for node in following:
        new_run.append(node)
    run.addnext(new_run)
    return new_run


def _split_t(run, t, offset: int):
    text = t.text or ''
    right = etree.Element(t.tag)
    right.text = text[offset:]
    right.set(XML_SPACE, 'preserve')
    t.text = text[:offset]
    t.set(XML_SPACE, 'preserve')
    t.addnext(right)
    _split_run_after(run, t)


def _ensure_boundary(para: Paragraph, doc: Document, pos: int) -> None:
    """Zadbaj, by w pozycji `pos` przebiegała granica runów (pętla do stabilizacji)."""
    while True:
        changed = False
        for item in para.items:
            if item.kind != 'text' or item.lock:
                continue
            content = _content_children(item.run)
            idx = content.index(item.t)
            if item.start < pos < item.end:
                _split_t(item.run, item.t, pos - item.start)
                changed = True
            elif item.end == pos and idx < len(content) - 1:
                changed = _split_run_after(item.run, item.t) is not None
            elif item.start == pos and idx > 0:
                changed = _split_run_after(item.run, content[idx - 1]) is not None
            if changed:
                break
        if not changed:
            return
        doc.rescan(para)


def _make_revision(tag: str, id_counter: IdCounter, author: str, date_str: str):
    el = etree.Element(f'{W}{tag}')
    el.set(f'{W}id', str(id_counter.next()))
    el.set(f'{W}author', author)
    el.set(f'{W}date', date_str)
    return el


def _apply_edit(para: Paragraph, doc: Document, s: int, e: int, ins: str,
                id_counter: IdCounter, author: str, date_str: str) -> None:
    if e == s:
        _ensure_boundary(para, doc, s)
        anchor, side = _insertion_anchor(para, s)
        rpr = anchor.run.find(f'{W}rPr')
        ins_el = make_ins(ins, rpr, id_counter=id_counter, author=author, date_str=date_str)
        (anchor.run.addnext if side == 'after' else anchor.run.addprevious)(ins_el)
        doc.rescan(para)
        return

    _ensure_boundary(para, doc, s)
    _ensure_boundary(para, doc, e)
    runs = []
    for item in para.items:
        if item.kind == 'text' and item.start < e and item.end > s and item.run not in runs:
            runs.append(item.run)
    first_rpr = runs[0].find(f'{W}rPr')
    first_rpr = copy.deepcopy(first_rpr) if first_rpr is not None else None
    for run in runs:
        for t in run.findall(f'{W}t'):
            t.tag = f'{W}delText'
            t.set(XML_SPACE, 'preserve')

    groups: list[list] = []
    for run in runs:
        if groups and groups[-1][-1].getnext() is run:
            groups[-1].append(run)
        else:
            groups.append([run])
    last_del = None
    for group in groups:
        del_el = _make_revision('del', id_counter, author, date_str)
        group[0].addprevious(del_el)
        for run in group:
            del_el.append(run)
        last_del = del_el
    if ins:
        last_del.addnext(make_ins(ins, first_rpr, id_counter=id_counter, author=author, date_str=date_str))
    doc.rescan(para)


def _capitalised_words(text: str) -> list[str]:
    return [w for w in re.findall(r'\w+', text) if w[:1].isupper()]


def apply_one(doc: Document, corr: dict, id_counter: IdCounter, author: str, date_str: str) -> dict:
    original, corrected = corr['original'], corr['corrected']
    result = {
        'original': original, 'corrected': corrected, 'note': corr.get('note', ''),
        'paragraph': corr.get('paragraph'), 'status': None, 'reason': '',
        'edits': [], 'deleted': '', 'inserted': '', 'warnings': [],
    }
    if original == corrected:
        result['status'] = STATUS_NOOP
        return result

    hits = _find_occurrences(doc, original)
    wanted = corr.get('paragraph')
    if wanted is not None:
        in_wanted = [h for h in hits if h[0].index == wanted]
        if not in_wanted:
            elsewhere = sorted({f'¶{h[0].index:03d}' for h in hits})
            result['status'] = STATUS_NOT_FOUND
            result['reason'] = f'brak w ¶{wanted:03d}' + (f' (jest w: {", ".join(elsewhere)})' if elsewhere else '')
            return result
        hits = in_wanted
    if not hits:
        result['status'] = STATUS_NOT_FOUND
        result['reason'] = 'nie znaleziono dokładnie takiego fragmentu (wielkość liter i znaki mają znaczenie)'
        return result
    if len(hits) > 1:
        where = ', '.join(sorted({f'¶{h[0].index:03d}' for h in hits}))
        result['status'] = STATUS_REJECTED
        result['reason'] = (f'niejednoznaczne: {len(hits)} wystąpień ({where}) — podaj "paragraph" '
                            'albo wydłuż "original"')
        return result

    para, offset = hits[0]
    result['paragraph'] = para.index
    text = para.text
    edits = [(offset + s, offset + e, ins) for s, e, ins in plan_edits(original, corrected)]
    for s, e, ins in edits:
        reason = _check_span(para, s, e, ins) or _check_content(text, s, e, ins)
        if reason:
            result['status'] = STATUS_REJECTED
            result['reason'] = reason
            return result

    result['edits'] = [{'deleted': text[s:e], 'inserted': ins} for s, e, ins in edits]
    result['deleted'] = ' … '.join(x['deleted'] for x in result['edits'] if x['deleted'])
    result['inserted'] = ' … '.join(x['inserted'] for x in result['edits'] if x['inserted'])
    if _capitalised_words(result['deleted']):
        result['warnings'].append('zmieniono wyraz z wielkiej litery — sprawdź, czy to nie nazwa własna')
    for s, e, ins in sorted(edits, reverse=True):
        _apply_edit(para, doc, s, e, ins, id_counter, author, date_str)
    result['status'] = STATUS_APPLIED
    return result


# ── API ────────────────────────────────────────────────────────────────────

def _id_counter_for(word_dir: str, root) -> IdCounter:
    roots = [root]
    for path in glob.glob(os.path.join(word_dir, '*.xml')):
        if os.path.basename(path) == 'document.xml':
            continue
        try:
            with open(path, 'rb') as f:
                roots.append(parse_xml(f.read()))
        except etree.XMLSyntaxError:
            continue
    return IdCounter(max_annotation_id(roots))


def apply_all(doc_path: str, corrections: list[dict], author: str, date_str: str,
              results: list | None = None) -> tuple[int, int]:
    """Nanieś korekty na rozpakowany word/document.xml. Zwraca (naniesione, wszystkie_nie-noop)."""
    corrections = validate_corrections(corrections)
    with open(doc_path, 'rb') as f:
        root = parse_xml(f.read())
    doc = Document(root)
    id_counter = _id_counter_for(os.path.dirname(doc_path), root)
    out = [apply_one(doc, corr, id_counter, author, date_str) for corr in corrections]
    etree.ElementTree(root).write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    if results is not None:
        results.extend(out)
    counted = [r for r in out if r['status'] != STATUS_NOOP]
    return sum(r['status'] == STATUS_APPLIED for r in counted), len(counted)


def apply_docx(docx_path: str, corrections: list[dict], output_path: str,
               author: str = 'Korektor AI', date_str: str | None = None) -> list[dict]:
    from docx_io import pack, unpack

    date_str = date_str or datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    results: list[dict] = []
    with tempfile.TemporaryDirectory() as tmpdir:
        unpack(docx_path, tmpdir)
        apply_all(os.path.join(tmpdir, 'word', 'document.xml'), corrections, author, date_str, results)
        pack(tmpdir, output_path, original_docx=docx_path)
    return results


def apply_correction(para, original, corrected, author, date_str, id_counter) -> bool:
    """Zgodność wstecz: korekta w pojedynczym akapicie (bez numeracji dokumentu)."""
    items, _ = scan_paragraph(para)
    paragraph = Paragraph(para, 1, (), items)
    holder = Document.__new__(Document)
    holder.root = para
    holder.paragraphs = [paragraph]
    result = apply_one(holder, {'original': original, 'corrected': corrected}, id_counter, author, date_str)
    return result['status'] == STATUS_APPLIED


# ── Raport ─────────────────────────────────────────────────────────────────

def _cell(text: str) -> str:
    return text.replace('|', '\\|').replace('\n', ' ')


def _describe(result: dict) -> str:
    parts = []
    for ed in result['edits']:
        if ed['deleted'] and ed['inserted']:
            parts.append(f'„{ed["deleted"]}” → „{ed["inserted"]}”')
        elif ed['inserted']:
            parts.append(f'wstawiono „{ed["inserted"]}”')
        else:
            parts.append(f'usunięto „{ed["deleted"]}”')
    return '; '.join(parts) or f'„{result["original"]}” → „{result["corrected"]}”'


LABELS = {
    STATUS_APPLIED: 'NANIESIONO', STATUS_REJECTED: 'ODRZUCONO',
    STATUS_NOT_FOUND: 'NIE ZNALEZIONO', STATUS_NOOP: 'BEZ ZMIANY',
}


def format_report(results: list[dict]) -> str:
    counted = [r for r in results if r['status'] != STATUS_NOOP]
    applied = sum(r['status'] == STATUS_APPLIED for r in counted)
    lines = [f'NANIESIONO: {applied} z {len(counted)}', '',
             '| # | ¶ | status | zmiana | uwaga / powód |', '|---|---|---|---|---|']
    for n, r in enumerate(results, 1):
        para = f'{r["paragraph"]:03d}' if isinstance(r['paragraph'], int) else '—'
        why = r['reason'] if r['status'] != STATUS_APPLIED else r['note']
        if r['warnings']:
            why = (why + ' ⚠ ' if why else '⚠ ') + '; '.join(r['warnings'])
        lines.append(f'| {n} | {para} | {LABELS[r["status"]]} | {_cell(_describe(r))} | {_cell(why)} |')
    for status in (STATUS_REJECTED, STATUS_NOT_FOUND):
        count = sum(r['status'] == status for r in results)
        if count:
            lines.append('')
            lines.append(f'{LABELS[status]}: {count} — tych poprawek NIE ma w pliku; popraw corrections.json albo zgłoś je w raporcie.')
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='Nanieś korekty z JSON na DOCX jako śledzone zmiany')
    parser.add_argument('docx', help='Plik wejściowy DOCX')
    parser.add_argument('corrections_json', help='Plik corrections.json')
    parser.add_argument('-o', '--output', default=None, help='Plik wyjściowy (domyślnie <wejście>_corrected.docx)')
    parser.add_argument('--author', default='Korektor AI', help='Autor śledzonych zmian')
    parser.add_argument('--date', default=None, help='Data śledzonych zmian (ISO)')
    parser.add_argument('--report', default=None, help='Zapisz wyniki per korekta do pliku JSON')
    args = parser.parse_args()

    output = args.output or '{}_corrected{}'.format(*os.path.splitext(args.docx))
    try:
        with open(args.corrections_json, encoding='utf-8') as f:
            corrections = validate_corrections(json.load(f))
    except (OSError, json.JSONDecodeError, CorrectionsError) as exc:
        print(f'BŁĄD corrections.json: {exc}', file=sys.stderr)
        sys.exit(2)

    try:
        results = apply_docx(args.docx, corrections, output, author=args.author, date_str=args.date)
    except Exception as exc:  # noqa: BLE001 — każdy błąd = brak pliku wyjściowego i kod 2
        if os.path.exists(output):
            os.remove(output)
        print(f'BŁĄD: {exc}', file=sys.stderr)
        sys.exit(2)

    print(format_report(results))
    print(f'\nPlik: {output}')
    if args.report:
        with open(args.report, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

    counted = [r for r in results if r['status'] != STATUS_NOOP]
    applied = sum(r['status'] == STATUS_APPLIED for r in counted)
    if counted and applied == 0:
        sys.exit(2)
    sys.exit(0 if applied == len(counted) else 1)


if __name__ == '__main__':
    logging.basicConfig(level=os.environ.get('LOG_LEVEL', 'WARNING').upper(), format='%(levelname)s: %(message)s')
    main()
