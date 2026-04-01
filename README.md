# Proofread DOCX — korekta z track changes

Korekta językowa pism procesowych (.docx) z nanoszeniem poprawek jako tracked changes (`w:del` / `w:ins`). Formatowanie dokumentu pozostaje nienaruszone.

## Instalacja

```bash
pip install lxml requests defusedxml
```

Python 3.10+. Brak zależności od zewnętrznych skryptów — unpack/pack wbudowany.

## Szybki start

```bash
# Automatyczna korekta (LanguageTool + reguły prawnicze)
python proofread_docx.py input.docx output.docx

# Wynik: output.docx z tracked changes, gotowy do review w MS Word
```

## Struktura modułów

```
proofreading/
├── proofread_docx.py      # Orkiestrator: unpack → check → apply → pack
├── apply_corrections.py   # Nanoszenie korekcji z JSON jako tracked changes
├── checkers.py            # Pluggable checkery (LanguageTool, regex, composite)
├── docx_io.py             # Unpack/pack .docx (zipfile, zero zew. zależności)
├── ooxml.py               # Wspólne helpery OOXML (namespace, run builder, ID counter)
├── merge_runs.py          # Scalanie sąsiednich runów z identycznym formatowaniem
├── verify_docx.py         # Weryfikacja integralności dokumentu po korekcie
└── tests/                 # 74 testy pytest
```

## Użycie

### Tryb 1: Automatyczny

```bash
python proofread_docx.py input.docx [output.docx]
```

Przepuszcza każdy paragraf przez pipeline checkerów, nanosi poprawki jako tracked changes. Domyślny output: `input_proofread.docx`.

### Tryb 2: Z gotowym JSON korekcji

Gdy korekcje generuje Claude lub inny proces:

```bash
# 1. Przygotuj corrections.json
cat corrections.json
```
```json
[
  {"original": "należnosci", "corrected": "należności", "note": "ortografia"},
  {"original": "od tego czy", "corrected": "od tego, czy", "note": "interpunkcja"}
]
```
```bash
# 2. Naniesienie na rozpakowany .docx
python apply_corrections.py unpacked_dir/ corrections.json --author "Korektor AI"
```

Pole `note` jest opcjonalne.

## Checkery

Architektura pozwala łączyć wiele źródeł korekty:

```
CompositeChecker
├── LanguageToolChecker   (API — ortografia, gramatyka, interpunkcja)
└── RegexChecker          (reguły regex — styl, formatowanie, prawo)
```

### Wbudowane reguły prawnicze

`polish_legal_rules()` zawiera:
- Podwójne spacje
- `w/w` → `ww.`
- Brak spacji po `art.`
- Brak przecinka przed `który/która/które`

### Dodawanie własnych reguł

```python
from checkers import RegexChecker, CompositeChecker, LanguageToolChecker

custom = RegexChecker()
custom.add_rule(r'\btzn\b', 'tzn.', 'Brak kropki po skrócie')
custom.add_rule(r'sąd rejonowy', 'Sąd Rejonowy', 'Wielka litera w nazwie sądu')

checker = CompositeChecker([LanguageToolChecker(), custom])
```

### Własny backend

Zaimplementuj `BaseChecker.check(text) -> list[Match]`:

```python
from checkers import BaseChecker

class MyChecker(BaseChecker):
    def check(self, text):
        # Match = {"offset": int, "length": int, "end": int,
        #          "replacement": str, "message": str,
        #          "rule_id": str, "category": str}
        return [...]
```

## Konfiguracja (env vars)

| Zmienna | Domyślna | Opis |
|---------|----------|------|
| `LANGUAGETOOL_URL` | `https://api.languagetool.org/v2/check` | URL API LanguageTool |
| `PROOFREAD_LANGUAGE` | `pl-PL` | Język korekty |
| `PROOFREAD_AUTHOR` | `Korektor` | Autor tracked changes w dokumencie |
| `PROOFREAD_REQUEST_DELAY` | `0.3` | Opóźnienie między requestami do API (s) |
| `LOG_LEVEL` | `INFO` | Poziom logowania (`DEBUG`, `WARNING`, itd.) |

Przykład z lokalnym LanguageTool Server:

```bash
LANGUAGETOOL_URL=http://localhost:8081/v2/check python proofread_docx.py input.docx
```

## Testy

```bash
pip install pytest
python -m pytest tests/ -v
```

74 testy pokrywające: ekstrakcję tekstu, manipulację XML, tracked changes, checkery, unpack/pack, merge runów.

## Scalanie runów (`merge_runs.py`)

Word dzieli tekst na wiele `<w:r>` z identycznym formatowaniem (np. po edycji, przy granicach stron). `merge_runs` scala je przed korektą, żeby LanguageTool widział pełne zdania.

```bash
# Użycie standalone (na rozpakowanym .docx):
python -c "from merge_runs import merge_runs; print(merge_runs('unpacked_dir/'))"
```

Co robi:
- Scala sąsiednie `<w:r>` z identycznym `<w:rPr>`
- Usuwa atrybuty `rsid` (metadane rewizji) i `proofErr` (markery spell-check)
- Stripuje `lastRenderedPageBreak` (rendering hint regenerowany przez Word)
- **Asercja integralności**: porównuje tekst przed/po — odmawia zapisu przy utracie >1%

## Weryfikacja dokumentu (`verify_docx.py`)

Porównanie oryginału z wersją po korekcie — wykrywa utratę tekstu, brakujące akapity, zmiany strukturalne.

```bash
python verify_docx.py oryginal.docx skorygowany.docx
python verify_docx.py oryginal.docx skorygowany.docx --strict   # warnings → errors
python verify_docx.py oryginal.docx skorygowany.docx --dump     # wypisz akapity
```

Sprawdza: liczbę akapitów, liczbę znaków (próg 1%/5%), diff per-paragraf, tabele, obrazy.

Exit codes: `0`=OK, `1`=WARNING, `2`=ERROR, `3`=FATAL.

## Gwarancje techniczne

- `w:rPr` (formatowanie runów) kopiowane z oryginału do tracked changes
- Operacje wyłącznie na warstwie tekstowej — brak ingerencji w style, nagłówki, stopki
- Case-insensitive fallback przy wyszukiwaniu fragmentów
- Korekcje przechodzące przez granice runów (multi-run spanning)
- Wiele korekcji w jednym runie (re-scan po każdej)
- Zachowanie kolejności ZIP memberów przy repackowaniu
- Asercja integralności tekstu w `merge_runs` (odmowa zapisu przy utracie danych)
