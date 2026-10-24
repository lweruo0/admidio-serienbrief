# -*- coding: utf-8 -*-
"""Prüft, ob zu jedem Datensatz einer CSV-Datei die erwartete PDF im Zielordner vorhanden ist.

Der Dateiname wird wie im Serienbrief-Client aus dem Muster gebildet (Platzhalter in
geschweiften Klammern, unzulässige Zeichen werden durch "_" ersetzt).

Beispiel:
    python check_pdfs.py daten.csv 2026-10-24-Rundschreiben
"""

import argparse
import csv
import io
import re
import sys
from pathlib import Path

DEFAULT_PATTERN = "rundschreiben-{Mitgliedsnummer}-{username}.pdf"
FORBIDDEN_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def read_text(path: Path) -> str:
    """Liest die Datei als UTF-8 (ohne BOM) mit Rückfall auf ISO-8859-1."""
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def detect_delimiter(text: str) -> str:
    """Bestimmt das Trennzeichen anhand der Kopfzeile; bei Gleichstand gewinnt das Semikolon."""
    first_line = text.split("\n", 1)[0]
    return max(["\t", ",", ";"], key=lambda candidate: (first_line.count(candidate), candidate == ";"))


def read_records(path: Path) -> list[dict[str, str]]:
    """Liest alle nicht leeren Datensätze als Dictionaries."""
    text = read_text(path)
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=detect_delimiter(text))
    rows = list(reader)
    if not rows:
        return []
    header = [name.strip() for name in rows[0]]
    records = []
    for row in rows[1:]:
        if all(not cell.strip() for cell in row):
            continue
        records.append(dict(zip(header, row)))
    return records


def render(pattern: str, values: dict[str, str]) -> str:
    """Ersetzt die Platzhalter des Musters durch bereinigte Werte."""
    result = pattern
    for name in dict.fromkeys(re.findall(r"\{([^{}]*)\}", pattern)):
        key = name.strip()
        if key in values:
            result = result.replace("{" + name + "}", FORBIDDEN_CHARS.sub("_", values[key].strip()))
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Prüft, ob zu jedem CSV-Datensatz eine PDF existiert.")
    parser.add_argument("csv_file", type=Path, help="CSV-Datenquelle")
    parser.add_argument("pdf_dir", type=Path, help="Ordner mit den erzeugten PDFs")
    parser.add_argument("--pattern", default=DEFAULT_PATTERN, help=f"Dateinamen-Muster (Standard: {DEFAULT_PATTERN})")
    args = parser.parse_args()

    if not args.csv_file.is_file():
        print(f"CSV-Datei nicht gefunden: {args.csv_file}")
        return 2
    if not args.pdf_dir.is_dir():
        print(f"Ordner nicht gefunden: {args.pdf_dir}")
        return 2

    records = read_records(args.csv_file)
    if not records:
        print("Die CSV enthält keine Datensätze.")
        return 2

    existing = {path.name.lower(): path.name for path in args.pdf_dir.iterdir() if path.is_file()}
    expected: dict[str, list[int]] = {}
    missing = []
    for number, record in enumerate(records, start=1):
        filename = render(args.pattern, record)
        expected.setdefault(filename.lower(), []).append(number)
        if filename.lower() not in existing:
            missing.append((number, filename))

    duplicates = {name: numbers for name, numbers in expected.items() if len(numbers) > 1}
    extra = sorted(name for key, name in existing.items() if key not in expected and key.endswith(".pdf"))

    print(f"{len(records)} Datensätze, {len(records) - len(missing)} PDFs vorhanden, {len(missing)} fehlen.")
    for number, filename in missing:
        print(f"  FEHLT  Datensatz {number}: {filename}")
    for name, numbers in duplicates.items():
        print(f"  DOPPELT  Datensätze {', '.join(map(str, numbers))} ergeben denselben Dateinamen: {name}")
    for name in extra:
        print(f"  ÜBRIG  PDF ohne Datensatz: {name}")

    return 0 if not missing and not duplicates else 1


if __name__ == "__main__":
    sys.exit(main())
