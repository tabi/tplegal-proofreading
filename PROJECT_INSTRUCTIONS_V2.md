# Project Instructions v2 — Zczytywacz (Claude-as-corrector)

## ZAKAZY (przeczytaj PRZED rozpoczęciem pracy)

- **NIE zmieniaj formatowania** dokumentu (czcionka, rozmiar, justowanie, nagłowki, stopki)
- **NIE dotykaj** sygnatur orzeczeń, numerów KRS/NIP/REGON, numerów faktur, dat, kwot
- **NIE zmieniaj** nazwisk, nazw firm, nazw miejscowości — chyba że masz 100% pewność co do literówki
- **NIE dodawaj** tekstu od siebie — korekta to TYLKO poprawianie istniejących bledow
- **NIE usuwaj** akapitów, tabel, list numerowanych
- **NIE koryguj** cytatów z ustaw, orzeczeń, umów (nawet jeśli zawierają błedy — to celowe cytaty)
- **NIE zmieniaj** skrótów prawniczych (art., ust., pkt, lit., zd., t.j., Dz.U.)

## Instalacja narzędzi

```bash
pip install git+https://github.com/tabi/tplegal-proofreading.git --break-system-packages -q
```

## Workflow korekty

### Krok 1: Ekstrakcja tekstu

```bash
extract-text input.docx > tekst.txt
```

Plik `tekst.txt` zawiera ponumerowane akapity w formacie `¶001: tekst...`

### Krok 2: Analiza i generowanie korekcji

Przeczytaj `tekst.txt` i wygeneruj plik `corrections.json`:

```json
[
  {
    "original": "tekst z bledem",
    "corrected": "tekst z błędem",
    "note": "brak polskiego znaku: e → ę"
  },
  {
    "original": "umwoa",
    "corrected": "umowa",
    "note": "literówka"
  }
]
```

**Zasady generowania korekcji:**

1. Pole `original` musi DOKŁADNIE odpowiadać tekstowi z dokumentu (case-sensitive)
2. Koryguj: ortografia, interpunkcja, literówki, fleksja, składnia
3. Nie koryguj: styl, kolejność słow, sygnonimy
4. Pole `note` — krotki opis rodzaju błędu (dla recenzenta)
5. Jedna korekta = jeden błąd. Nie łącz wielu bledow w jedną korektę
6. `original` powinien zawierać minimum kontekstu potrzebnego do jednoznacznego dopasowania (zwykle 3-5 słów wokół błędu)

### Krok 3: Naniesienie korekcji

```bash
apply-corrections input.docx corrections.json -o output.docx
```

Opcje:
- `--author "Imię Nazwisko"` — autor tracked changes (domyślnie: "Korektor AI")
- `--date "2026-04-01T12:00:00Z"` — data tracked changes (domyślnie: teraz UTC)

Kody wyjścia: 0 = wszystkie naniesione, 1 = częściowo (niektóre nie znalezione), 2 = błąd

### Krok 4: Weryfikacja integralności

```bash
verify-docx input.docx output.docx
```

Sprawdza czy korekta nie uszkodziła dokumentu (liczba akapitów, znaków, tabele, obrazy).

**WAŻNE:** Tracked changes w Word wymagają 2 kliknięć per korekta (Accept/Reject osobno dla usunięcia i wstawienia). To normalne zachowanie, nie bug.

## Czego NIE robić

- Nie uruchamiaj apply-corrections na pliku który już ma tracked changes z poprzedniego przebiegu
- Nie edytuj output.docx ręcznie po apply-corrections — to zepsuje weryfikację
- Nie generuj korekcji z pamięci — ZAWSZE bazuj na tekście z extract-text
- Nie ignoruj exit code 1 z apply-corrections — sprawdź logi, popraw corrections.json
- Nie pomijaj verify-docx — to jedyne zabezpieczenie przed uszkodzeniem dokumentu
