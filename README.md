# tplegal-proofreading — korekta DOCX z track changes

Korekta językowa pism (.docx): Claude czyta tekst i wskazuje poprawki, pakiet nanosi je jako śledzone zmiany (`w:del` / `w:ins`) i sprawdza, że nic poza nimi się nie zmieniło.

## Instalacja

```bash
pip install git+https://github.com/tabi/tplegal-proofreading.git
```

Python 3.10+, zależności: `lxml`, `defusedxml`.

## Trzy komendy

```bash
extract-text input.docx > tekst.txt
apply-corrections input.docx corrections.json -o output.docx
verify-docx input.docx output.docx
```

### `extract-text`

Numerowane akapity treści głównej (`¶001: …`) — ten sam tekst i numeracja, na których pracuje `apply-corrections`. Znaczniki elementów nietekstowych: `\t` tabulator, `↵` złamanie wiersza, `[^N]` przypis, `[obraz]`, `[wzór]`, `□` symbol. Przypisy, nagłówki, stopki i pola tekstowe nie są czytane.

### `apply-corrections`

```json
[
  {"original": "od tego czy", "corrected": "od tego, czy", "note": "interpunkcja", "paragraph": 12}
]
```

- `original` — dokładny fragment (wielkość liter ma znaczenie, brak fallbacku),
- `paragraph` — opcjonalny numer ¶; wymagany, gdy `original` występuje w dokumencie więcej niż raz.

Korekta jest **odrzucana z powodem** (nic nie trafia do pliku), gdy:

- `original` jest niejednoznaczny albo go nie ma,
- zmiana obejmowałaby tabulator, złamanie wiersza, przypis, obraz lub pole Worda,
- zmiana wchodzi w istniejącą śledzoną zmianę (cudzą albo z poprzedniego przebiegu),
- poprawiany fragment ma mieszane formatowanie albo przechodzi przez granicę hiperłącza,
- zmienia cyfry albo dotyka znaków `& @ / \ _` (sygnatury, numery, daty, kwoty, nazwy, adresy).

Śledzona zmiana obejmuje tylko wyrazy/znaki faktycznie różne (`od tego czy` → `od tego, czy` = samo wstawienie przecinka). Na stdout: tabela wszystkich korekt ze statusem. `--report plik.json` zapisuje to samo jako JSON.

Kody wyjścia: `0` wszystkie naniesione, `1` część odrzucona/nieznaleziona, `2` błąd albo nic nie naniesiono.

### `verify-docx`

Bramka przed zwrotem pliku. PASS tylko wtedy, gdy:

1. wszystkie pliki w ZIP poza `word/document.xml` są identyczne bajt w bajt,
2. `document.xml` po odrzuceniu nowych śledzonych zmian jest strukturalnie identyczny z wejściem (tekst, formatowanie, tabulatory, przypisy, obrazy, pola, tabele; pomijane są tylko atrybuty `w:rsid*`),
3. znaczniki nowych zmian są poprawne (unikalne `w:id`, brak zagnieżdżeń).

Wypisuje **pełną listę zmian odczytaną z pliku** (¶, było, jest, kontekst) — niezależnie od tego, co raportuje korektor. `--strict-text` przyjmowane dla zgodności (kontrola ścisła jest zawsze włączona).

Kody wyjścia: `0` PASS, `2` FAIL (nie zwracaj pliku), `3` plik nieczytelny.

## Moduły

| Plik | Rola |
|---|---|
| `docmodel.py` | wspólny model tekstu akapitu (jeden czytnik dla wszystkich komend) |
| `extract_text.py` | CLI `extract-text` |
| `apply_corrections.py` | CLI `apply-corrections` |
| `verify_docx.py` | CLI `verify-docx` |
| `ooxml.py`, `minidom_helpers.py` | helpery OOXML |
| `docx_io.py` | unpack/pack .docx z zachowaniem kolejności plików |

## Testy

```bash
python3 -m pytest tests -q
```

`tests/test_v23_bezpieczne_nanoszenie.py` — regresje z audytu 24.09.2026 (ciche usuwanie tabulatorów/przypisów, przelewanie formatowania, naniesienie w złe miejsce, rozjazd tekstu przy drugim przebiegu i hiperłączach).

## Historia

- **2.3.0** (24.09.2026) — wspólny docmodel, minimalne śledzone zmiany, blokady, `verify-docx` ze ścisłą kontrolą i listą zmian. Wersja „2.2” opisywana w Outline nigdy nie trafiła do repo.
- **2.0.0** (01.04.2026) — Claude-as-corrector, 3 CLI.
- **1.x** (03.2026) — LanguageTool, wycofany (psuł nazwy własne i sygnatury).
