#!/usr/bin/env python3
"""Weryfikacja DOCX po korekcie — bramka przed zwrotem pliku.

Sprawdza, że JEDYNE różnice między input.docx a output.docx to nowe śledzone
zmiany (w:ins/w:del), i wypisuje ich PEŁNĄ listę odczytaną z samego pliku
(nie z raportu korektora).

Kontrole (zawsze włączone; flaga --strict-text zostaje dla zgodności):
  1. Wszystkie pliki w ZIP poza word/document.xml — identyczne bajt w bajt.
  2. word/document.xml po odrzuceniu NOWYCH śledzonych zmian == input
     strukturalnie: tekst, formatowanie, tabulatory, przypisy, obrazy, pola,
     tabele, sekcje (po scaleniu runów o identycznym formatowaniu).
  3. Poprawność znaczników nowych zmian: unikalne w:id, w:del bez w:t,
     w:ins bez w:delText, brak zagnieżdżeń.

Usage:
    verify-docx input.docx output.docx [--dump]

Exit codes: 0 = PASS, 2 = FAIL (nie zwracaj pliku), 3 = FATAL (plik nieczytelny).
"""

import argparse
import copy
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import defusedxml.minidom as minidom_mod
from lxml import etree

from docmodel import REVISION_TAGS, W, XML_SPACE, Document, body_paragraph_elements, parse_xml
from minidom_helpers import extract_paragraph_text as _extract_paragraph_text
from minidom_helpers import find_elements as _find_elements

DOC = 'word/document.xml'
CONTEXT = 40


# ── Statystyki (legacy, --dump) ───────────────────────────────────────────

@dataclass
class DocStats:
    paragraphs: list[str] = field(default_factory=list)
    total_chars: int = 0
    table_count: int = 0
    image_count: int = 0
    header_footer_count: int = 0


def extract_stats(docx_path: str, mode: str = "visible") -> DocStats:
    """Tekst akapitów i liczniki struktury. mode: 'visible' | 'original'."""
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")

    stats = DocStats()
    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        stats.header_footer_count = sum(
            1 for n in names if n.startswith("word/header") or n.startswith("word/footer")
        )
        if DOC not in names:
            raise ValueError(f"No word/document.xml in {docx_path}")
        root = minidom_mod.parseString(zf.read(DOC).decode("utf-8")).documentElement
        for p_elem in _find_elements(root, "p"):
            text = _extract_paragraph_text(p_elem, mode=mode)
            stats.paragraphs.append(text)
            stats.total_chars += len(text)
        stats.table_count = len(_find_elements(root, "tbl"))
        stats.image_count = len(_find_elements(root, "drawing")) + len(_find_elements(root, "pict"))
    return stats


# ── Normalizacja i odrzucanie nowych zmian ────────────────────────────────

def _revision_ids(root) -> set[str]:
    return {el.get(f'{W}id') for tag in REVISION_TAGS for el in root.iter(tag)}


def _reject_new(root, old_ids: set[str]) -> None:
    """Odrzuć (in place) śledzone zmiany, których nie było w pliku wejściowym."""
    for el in list(root.iter(f'{W}ins')):
        if el.get(f'{W}id') not in old_ids:
            el.getparent().remove(el)
    for el in list(root.iter(f'{W}del')):
        if el.get(f'{W}id') in old_ids:
            continue
        for dt in el.iter(f'{W}delText'):
            dt.tag = f'{W}t'
        parent = el.getparent()
        idx = parent.index(el)
        for offset, child in enumerate(list(el)):
            parent.insert(idx + offset, child)
        parent.remove(el)


def _rpr_bytes(run) -> bytes:
    rpr = run.find(f'{W}rPr')
    return etree.tostring(rpr, method='c14n') if rpr is not None else b''


def _normalize(root) -> None:
    """Scal runy/w:t, które różnią się tylko podziałem (skutek dzielenia runów przy korekcie).

    Atrybuty w:rsid* (identyfikatory sesji edycji Worda) są pomijane — nie wpływają
    na treść ani wygląd, a Word zmienia je przy każdym zapisie.
    """
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        for key in [k for k in el.attrib if k.startswith(f'{W}rsid')]:
            del el.attrib[key]
        if el.tag not in (f'{W}t', f'{W}delText', f'{W}instrText'):
            if el.text is not None and not el.text.strip():
                el.text = None
        if el.tail is not None and not el.tail.strip():
            el.tail = None
    for parent in list(root.iter()):
        child = parent[0] if len(parent) else None
        while child is not None:
            nxt = child.getnext()
            if (nxt is not None and child.tag == f'{W}r' and nxt.tag == f'{W}r'
                    and dict(child.attrib) == dict(nxt.attrib) and _rpr_bytes(child) == _rpr_bytes(nxt)):
                for sub in list(nxt):
                    if sub.tag != f'{W}rPr':
                        child.append(sub)
                parent.remove(nxt)
                continue
            child = nxt
    for run in root.iter(f'{W}r'):
        prev = None
        for sub in list(run):
            if sub.tag == f'{W}t':
                sub.attrib.pop(XML_SPACE, None)
                if not sub.text:
                    run.remove(sub)
                    continue
                if prev is not None and prev.tag == f'{W}t' and prev.getnext() is sub:
                    prev.text = (prev.text or '') + sub.text
                    run.remove(sub)
                    continue
            prev = sub


def _canon(el) -> bytes:
    return etree.tostring(el, method='c14n')


def _plain_text(p) -> str:
    return ''.join(t.text or '' for t in p.iter(f'{W}t'))


# ── Kontrola znaczników nowych zmian ─────────────────────────────────────

def _markup_errors(root, old_ids: set[str]) -> list[str]:
    errors = []
    ids = [el.get(f'{W}id') for tag in REVISION_TAGS for el in root.iter(tag)]
    dup = sorted(i for i, n in Counter(ids).items() if n > 1)
    if dup:
        errors.append(f'zdublowane w:id śledzonych zmian: {", ".join(dup[:10])}')
    for tag in (f'{W}ins', f'{W}del'):
        for el in root.iter(tag):
            if el.get(f'{W}id') in old_ids:
                continue
            anc = el.getparent()
            while anc is not None:
                if anc.tag in REVISION_TAGS:
                    errors.append(f'zagnieżdżona śledzona zmiana w:id={el.get(f"{W}id")}')
                    break
                anc = anc.getparent()
            for child in el:
                if child.tag != f'{W}r':
                    errors.append(f'w:{tag.split("}")[1]} w:id={el.get(f"{W}id")} zawiera {child.tag.split("}")[-1]} zamiast runów')
                    continue
                bad = f'{W}t' if tag == f'{W}del' else f'{W}delText'
                if child.find(bad) is not None:
                    errors.append(f'w:{tag.split("}")[1]} w:id={el.get(f"{W}id")} zawiera {bad.split("}")[1]}')
    return errors


# ── Lista zmian ───────────────────────────────────────────────────────────

def _change_pieces(p, old_ids: set[str]) -> list[tuple[str, str, object]]:
    """Kawałki akapitu w kolejności: ('plain'|'del'|'ins', tekst, element zmiany)."""
    pieces = []

    def walk(node, mode, rev):
        for child in node:
            tag = child.tag
            if not isinstance(tag, str):
                continue
            if tag in (f'{W}ins', f'{W}del', f'{W}moveFrom', f'{W}moveTo'):
                if child.get(f'{W}id') in old_ids:
                    if tag in (f'{W}ins', f'{W}moveTo'):
                        walk(child, mode, rev)
                    continue
                walk(child, 'ins' if tag == f'{W}ins' else 'del', child)
            elif tag in (f'{W}t', f'{W}delText'):
                pieces.append((mode, child.text or '', rev))
            elif tag in (f'{W}drawing', f'{W}pict', f'{W}txbxContent'):
                continue
            else:
                walk(child, mode, rev)

    walk(p, 'plain', None)
    return pieces


def list_changes(out_root, in_root, old_ids: set[str]) -> list[dict]:
    in_numbers = [p.index for p in Document(in_root).paragraphs]
    out_paras = body_paragraph_elements(out_root)
    changes = []
    for k, p in enumerate(out_paras):
        pieces = _change_pieces(p, old_ids)
        if not any(m != 'plain' for m, _, _ in pieces):
            continue
        number = in_numbers[k] if k < len(in_numbers) else None
        i = 0
        while i < len(pieces):
            if pieces[i][0] == 'plain':
                i += 1
                continue
            # Jedna zmiana = jeden w:del/w:ins albo para w:del + bezpośrednio po nim w:ins.
            j = i
            deleted, inserted = [], []
            revs = []
            while j < len(pieces) and pieces[j][0] != 'plain':
                mode, text, rev = pieces[j]
                if revs and rev is not revs[-1]:
                    pair = (revs[-1].tag == f'{W}del' and rev.tag == f'{W}ins'
                            and len(revs) == 1 and revs[-1].getnext() is rev)
                    if not pair:
                        break
                (deleted if mode == 'del' else inserted).append(text)
                if rev not in revs:
                    revs.append(rev)
                j += 1
            before = ''.join(t for m, t, _ in pieces[:i] if m != 'del')[-CONTEXT:]
            after = ''.join(t for m, t, _ in pieces[j:] if m != 'del')[:CONTEXT]
            changes.append({
                'paragraph': number, 'deleted': ''.join(deleted), 'inserted': ''.join(inserted),
                'before': before, 'after': after,
                'author': revs[0].get(f'{W}author', '') if revs else '',
            })
            i = j
    return changes


# ── Weryfikacja ───────────────────────────────────────────────────────────

def verify(original_path: str, corrected_path: str) -> dict:
    """{'ok': bool, 'errors': [...], 'changes': [...], 'preexisting': int}. Wyjątek = plik nieczytelny."""
    errors: list[str] = []
    with zipfile.ZipFile(original_path) as za, zipfile.ZipFile(corrected_path) as zb:
        # Wpisy katalogów (np. "word/") nie niosą treści, a pack() ich nie odtwarza.
        names_a = {n for n in za.namelist() if not n.endswith('/')}
        names_b = {n for n in zb.namelist() if not n.endswith('/')}
        for name in sorted(names_a - names_b):
            errors.append(f'brak pliku w ZIP: {name}')
        for name in sorted(names_b - names_a):
            errors.append(f'nowy plik w ZIP: {name}')
        for name in sorted(names_a & names_b):
            if name != DOC and za.read(name) != zb.read(name):
                errors.append(f'zmieniony plik w ZIP (poza treścią główną): {name}')
        in_root = parse_xml(za.read(DOC))
        out_root = parse_xml(zb.read(DOC))

    old_ids = _revision_ids(in_root)
    errors.extend(_markup_errors(out_root, old_ids))

    rejected = copy.deepcopy(out_root)
    _reject_new(rejected, old_ids)
    baseline = copy.deepcopy(in_root)
    _normalize(rejected)
    _normalize(baseline)

    if _canon(rejected) != _canon(baseline):
        pa, pb = body_paragraph_elements(baseline), body_paragraph_elements(rejected)
        if len(pa) != len(pb):
            errors.append(f'zmieniona liczba akapitów: {len(pa)} → {len(pb)}')
        else:
            found = 0
            for k, (a, b) in enumerate(zip(pa, pb)):
                if _canon(a) == _canon(b):
                    continue
                found += 1
                ta, tb = _plain_text(a), _plain_text(b)
                if ta != tb:
                    errors.append(f'akapit {k + 1}: cicha zmiana tekstu poza śledzonymi zmianami: „{ta[:150]}” → „{tb[:150]}”')
                else:
                    errors.append(f'akapit {k + 1}: cicha zmiana formatowania/struktury (np. pogrubienie, tabulator, przypis) w „{ta[:150]}”')
            if not found:
                errors.append('cicha zmiana poza akapitami (właściwości tabeli, sekcji albo dokumentu)')

    changes = list_changes(out_root, in_root, old_ids)
    return {'ok': not errors, 'errors': errors, 'changes': changes, 'preexisting': len(old_ids)}


def _cell(text: str) -> str:
    return text.replace('|', '\\|').replace('\n', ' ')


def format_verify_report(report: dict) -> str:
    lines = [f'VERIFY: {"PASS" if report["ok"] else "FAIL — NIE ZWRACAJ PLIKU"}']
    if report['errors']:
        lines.append(f'BŁĘDY ({len(report["errors"])}):')
        lines.extend(f'  - {e}' for e in report['errors'])
    if report['preexisting']:
        lines.append(f'Śledzone zmiany obecne już w pliku wejściowym: {report["preexisting"]} (nietknięte, poza listą).')
    lines.append('')
    lines.append(f'ZMIANY W DOKUMENCIE: {len(report["changes"])}')
    if report['changes']:
        lines.append('| # | ¶ | było | jest | kontekst |')
        lines.append('|---|---|---|---|---|')
        for n, c in enumerate(report['changes'], 1):
            para = f'{c["paragraph"]:03d}' if c['paragraph'] else '—'
            ctx = f'…{c["before"]}⟦{c["deleted"]}→{c["inserted"]}⟧{c["after"]}…'
            lines.append(f'| {n} | {para} | {_cell(c["deleted"]) or "—"} | {_cell(c["inserted"]) or "—"} | {_cell(ctx)} |')
    return '\n'.join(lines)


def compare(original_path: str, corrected_path: str, strict: bool = False, dump: bool = False, quiet: bool = False) -> int:
    """Zwraca kod wyjścia: 0 = PASS, 2 = FAIL, 3 = FATAL."""
    try:
        report = verify(original_path, corrected_path)
    except Exception as e:  # noqa: BLE001 — nieczytelny plik = FATAL
        print(f"FATAL: nie można odczytać pliku: {e}", file=sys.stderr)
        return 3
    if not quiet:
        print(format_verify_report(report))
        if dump:
            for label, path, mode in (('WEJŚCIE', original_path, 'visible'), ('WYJŚCIE (po odrzuceniu zmian)', corrected_path, 'original')):
                print(f'\n--- {label} ---')
                for i, p in enumerate(extract_stats(path, mode=mode).paragraphs):
                    if p.strip():
                        print(f'  {i:03d}: {p[:120]}')
    return 0 if report['ok'] else 2


def main():
    parser = argparse.ArgumentParser(description="Weryfikacja DOCX po korekcie + pełna lista zmian")
    parser.add_argument("original", help="Plik wejściowy DOCX (przed korektą)")
    parser.add_argument("corrected", help="Plik wyjściowy DOCX (po korekcie)")
    parser.add_argument("--strict-text", action="store_true", help="Zgodność wstecz — kontrola ścisła jest zawsze włączona")
    parser.add_argument("--strict", action="store_true", help="Zgodność wstecz — bez działania")
    parser.add_argument("--dump", action="store_true", help="Wypisz akapity obu plików")
    args = parser.parse_args()
    sys.exit(compare(args.original, args.corrected, dump=args.dump))


if __name__ == "__main__":
    main()
