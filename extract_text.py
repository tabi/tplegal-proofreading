#!/usr/bin/env python3
"""Extract plain text from a DOCX file for proofreading.

Numerowane akapity treści głównej (¶001: …) — ten sam tekst i ta sama numeracja,
na których pracuje apply-corrections (wspólny docmodel). Widok = tak, jak Word
wyświetla dokument: bez tekstu usuniętego (w:del), z tekstem wstawionym (w:ins)
i hiperłączami.

Elementy nietekstowe są widoczne jako znaczniki, których NIE wolno zmieniać:
  \\t tabulator, ↵ złamanie wiersza, [^N] odnośnik przypisu, [obraz], [wzór], □ symbol, ¹ ₂ ^(…) indeks górny/dolny.

Usage:
    extract-text input.docx > tekst.txt
"""

import argparse
import sys
import zipfile
from pathlib import Path

from docmodel import Document, parse_xml

LEGEND = (
    'Zakres: treść główna. Przypisy, nagłówki, stopki i pola tekstowe NIE są czytane.\n'
    'Znaczniki (nie zmieniaj ich): \\t tabulator, ↵ złamanie wiersza, [^N] przypis, [obraz], [wzór], □ symbol, ¹ ₂ ^(…) indeks górny/dolny.'
)


def extract_text(docx_path: str) -> list[str]:
    """Extract numbered paragraph texts from a DOCX file.

    Returns list of strings like '¶001: Tekst akapitu...'
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")

    with zipfile.ZipFile(path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            raise ValueError(f"No word/document.xml in {docx_path}")
        root = parse_xml(zf.read("word/document.xml"))

    return [f"¶{p.index:03d}: {p.text}" for p in Document(root).numbered()]


def main():
    parser = argparse.ArgumentParser(
        description="Extract plain text from a DOCX file for proofreading"
    )
    parser.add_argument("docx", help="Input DOCX file")
    args = parser.parse_args()

    try:
        lines = extract_text(args.docx)
    except (FileNotFoundError, ValueError, zipfile.BadZipFile) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(2)

    for line in lines:
        print(line)
    print(f"{LEGEND}\nAkapitów: {len(lines)}", file=sys.stderr)


if __name__ == "__main__":
    main()
