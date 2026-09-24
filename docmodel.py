"""Wspólny model tekstu akapitu — JEDNO źródło prawdy dla extract-text,
apply-corrections i verify-docx.

v2.0 miało dwa niezależne czytniki (minidom w extract-text, bezpośrednie w:r
w apply-corrections), które widziały inny tekst: hiperlinki i treść w:ins były
widoczne dla modelu, a niewidoczne dla narzędzia nanoszącego. Tu jest jeden.

Akapit = lista elementów (Item) w kolejności dokumentu:
  - 'text'  — fragment z w:t (można korygować, chyba że `lock`),
  - 'token' — element nietekstowy (tabulator, złamanie wiersza, przypis, obraz,
              znacznik pola). Ma tekst wyświetlany modelowi, ale NIGDY nie
              może znaleźć się wewnątrz korygowanego fragmentu.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'
MC_ALT = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
M_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

REVISION_TAGS = (f'{W}ins', f'{W}del', f'{W}moveFrom', f'{W}moveTo')

LOCK_REVISION = 'fragment leży w istniejącej śledzonej zmianie (cudzej albo z poprzedniego przebiegu)'
LOCK_FIELD = 'fragment jest wynikiem pola Worda (np. DOCPROPERTY, spis treści) — Word nadpisze go przy aktualizacji'
LOCK_FORMAT_CHANGE = 'fragment ma śledzoną zmianę formatowania (w:rPrChange)'

# Akapity wewnątrz obiektów pływających — nie korygujemy (pole tekstowe bywa
# zapisane dwukrotnie: mc:Choice + mc:Fallback).
_EXCLUDED_ANCESTORS = {
    f'{W}drawing', f'{W}pict', f'{W}object', f'{W}txbxContent', MC_ALT,
}

# Znaczniki na poziomie akapitu bez treści widocznej — pomijane.
_IGNORABLE_P_CHILDREN = {
    f'{W}pPr', f'{W}bookmarkStart', f'{W}bookmarkEnd', f'{W}proofErr',
    f'{W}commentRangeStart', f'{W}commentRangeEnd', f'{W}permStart', f'{W}permEnd',
    f'{W}moveFromRangeStart', f'{W}moveFromRangeEnd', f'{W}moveToRangeStart', f'{W}moveToRangeEnd',
    f'{W}customXmlInsRangeStart', f'{W}customXmlInsRangeEnd',
    f'{W}customXmlDelRangeStart', f'{W}customXmlDelRangeEnd',
}
_TRANSPARENT_CONTAINERS = {f'{W}hyperlink', f'{W}smartTag', f'{W}customXml', f'{W}dir', f'{W}bdo'}

# Dzieci w:r bez treści widocznej — pomijane (nie blokują korekty).
_IGNORABLE_RUN_CHILDREN = {f'{W}rPr', f'{W}lastRenderedPageBreak', f'{W}instrText', f'{W}delInstrText', f'{W}delText'}

# Cechy w:rPr bez wpływu na wygląd — ignorowane przy porównaniu formatowania.
_RPR_NONVISUAL = {f'{W}lang', f'{W}noProof', f'{W}rFonts'}


def xml_parser() -> etree.XMLParser:
    """Parser bez rozwijania encji i bez sieci (plik .docx to niezaufane wejście)."""
    return etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False, huge_tree=True)


def parse_xml(data: bytes) -> etree._Element:
    return etree.fromstring(data, parser=xml_parser())


@dataclass
class Item:
    kind: str                       # 'text' | 'token'
    text: str                       # tekst widoczny (dla tokenu: jego reprezentacja)
    start: int = 0
    end: int = 0
    run: etree._Element | None = None
    t: etree._Element | None = None
    container: etree._Element | None = None
    lock: str | None = None
    name: str = ''                  # dla tokenu: nazwa elementu (do komunikatów)


@dataclass
class Paragraph:
    elem: etree._Element
    index: int                      # numer ¶ (1..n); 0 = akapit pusty, nienumerowany
    field_stack_in: tuple = ()
    items: list[Item] = field(default_factory=list)

    @property
    def text(self) -> str:
        return ''.join(i.text for i in self.items)


def rpr_signature(rpr: etree._Element | None) -> bytes:
    """Kanoniczny zapis formatowania runu bez cech niewizualnych."""
    if rpr is None:
        return b''
    clone = etree.Element(rpr.tag)
    for child in rpr:
        if child.tag in _RPR_NONVISUAL or not isinstance(child.tag, str):
            continue
        clone.append(etree.fromstring(etree.tostring(child)))
    return etree.tostring(clone, method='c14n')


def _local(tag) -> str:
    return tag.split('}', 1)[-1] if isinstance(tag, str) else ''


def _is_hidden(rpr: etree._Element | None) -> bool:
    if rpr is None:
        return False
    for tag in (f'{W}vanish', f'{W}webHidden', f'{W}specVanish'):
        el = rpr.find(tag)
        if el is not None and el.get(f'{W}val', 'true') not in ('0', 'false', 'off'):
            return True
    return False


class _Walker:
    def __init__(self, field_stack: list[str]):
        self.items: list[Item] = []
        self.fields = field_stack

    def token(self, text: str, name: str, elem=None, container=None):
        self.items.append(Item('token', text, run=elem, container=container, name=name))

    def walk(self, node, container, lock):
        for child in node:
            tag = child.tag
            if not isinstance(tag, str):
                continue
            if tag == f'{W}r':
                self.run(child, container, lock)
            elif tag in (f'{W}del', f'{W}moveFrom'):
                continue
            elif tag in (f'{W}ins', f'{W}moveTo'):
                self.walk(child, child, LOCK_REVISION)
            elif tag in _TRANSPARENT_CONTAINERS:
                self.walk(child, child, lock)
            elif tag == f'{W}sdt':
                content = child.find(f'{W}sdtContent')
                if content is not None:
                    self.walk(content, content, lock)
            elif tag == f'{W}fldSimple':
                self.walk(child, child, LOCK_FIELD)
            elif tag in _IGNORABLE_P_CHILDREN:
                continue
            elif tag.startswith(M_NS):
                self.token('[wzór]', _local(tag), child, container)
            else:
                self.token('', _local(tag), child, container)

    def run(self, r, container, lock):
        rpr = r.find(f'{W}rPr')
        if _is_hidden(rpr):
            self.token('', 'ukryty tekst', r, container)
            return
        run_lock = lock
        if run_lock is None and rpr is not None and rpr.find(f'{W}rPrChange') is not None:
            run_lock = LOCK_FORMAT_CHANGE
        for c in r:
            tag = c.tag
            if not isinstance(tag, str) or tag in _IGNORABLE_RUN_CHILDREN:
                continue
            if tag == f'{W}t':
                item_lock = run_lock or (LOCK_FIELD if self.fields else None)
                self.items.append(Item('text', c.text or '', run=r, t=c, container=container, lock=item_lock))
            elif tag in (f'{W}tab', f'{W}ptab'):
                self.token('\t', 'tabulator', r, container)
            elif tag in (f'{W}br', f'{W}cr'):
                self.token('↵', 'złamanie wiersza', r, container)
            elif tag == f'{W}noBreakHyphen':
                self.token('-', 'łącznik nierozdzielający', r, container)
            elif tag == f'{W}softHyphen':
                self.token('', 'miękki łącznik', r, container)
            elif tag == f'{W}sym':
                self.token('□', 'symbol', r, container)
            elif tag == f'{W}footnoteReference':
                self.token(f'[^{c.get(f"{W}id", "")}]', 'przypis', r, container)
            elif tag == f'{W}endnoteReference':
                self.token(f'[^e{c.get(f"{W}id", "")}]', 'przypis końcowy', r, container)
            elif tag == f'{W}fldChar':
                kind = c.get(f'{W}fldCharType')
                if kind == 'begin':
                    self.fields.append('instr')
                elif kind == 'separate' and self.fields:
                    self.fields[-1] = 'result'
                elif kind == 'end' and self.fields:
                    self.fields.pop()
                self.token('', 'znacznik pola', r, container)
            elif tag in (f'{W}drawing', f'{W}pict', f'{W}object') or tag == MC_ALT:
                self.token('[obraz]', 'obraz/obiekt', r, container)
            else:
                self.token('', _local(tag), r, container)


def scan_paragraph(p: etree._Element, field_stack_in: tuple = ()) -> tuple[list[Item], tuple]:
    """Zbuduj listę elementów akapitu. Zwraca (items, stan pól po akapicie)."""
    walker = _Walker(list(field_stack_in))
    walker.walk(p, p, None)
    pos = 0
    for it in walker.items:
        it.start = pos
        pos += len(it.text)
        it.end = pos
    return walker.items, tuple(walker.fields)


def _excluded(p: etree._Element, stop: etree._Element) -> bool:
    node = p.getparent()
    while node is not None and node is not stop:
        if node.tag in _EXCLUDED_ANCESTORS:
            return True
        node = node.getparent()
    return False


def body_paragraph_elements(root: etree._Element) -> list[etree._Element]:
    body = root.find(f'{W}body')
    if body is None:
        body = root
    return [p for p in body.iter(f'{W}p') if not _excluded(p, body)]


class Document:
    """Akapity treści głównej z numeracją ¶ zgodną z extract-text."""

    def __init__(self, root: etree._Element):
        self.root = root
        self.paragraphs: list[Paragraph] = []
        stack: tuple = ()
        n = 0
        for p in body_paragraph_elements(root):
            items, stack_out = scan_paragraph(p, stack)
            text = ''.join(i.text for i in items)
            index = 0
            if text.strip():
                n += 1
                index = n
            self.paragraphs.append(Paragraph(p, index, stack, items))
            stack = stack_out

    def numbered(self) -> list[Paragraph]:
        return [p for p in self.paragraphs if p.index]

    def rescan(self, para: Paragraph) -> None:
        para.items, _ = scan_paragraph(para.elem, para.field_stack_in)


def max_annotation_id(roots) -> int:
    """Największe liczbowe w:id we wszystkich podanych drzewach."""
    best = 0
    for root in roots:
        for el in root.iter():
            val = el.get(f'{W}id') if isinstance(el.tag, str) else None
            if val is not None and val.lstrip('-').isdigit():
                best = max(best, int(val))
    return best
