#!/usr/bin/env python3
"""Converti il CSV posizioni KiCad nel formato XLSX del modello CPL JLCPCB.

Dipendenza: python3 -m pip install openpyxl
Uso: python3 kicad_to_jlcpcb.py scheda-all-pos.csv [-o scheda-CPL.xlsx]

Le coordinate KiCad sono considerate mm, salvo --units in. Non vengono
specchiati assi, spostata l'origine o applicate correzioni ai footprint.
"""

import argparse
import csv
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
import sys
import tempfile

HEADERS = ("Designator", "Mid X", "Mid Y", "Layer", "Rotation")
REQUIRED = ("Ref", "PosX", "PosY", "Rot", "Side")


def number(value, line, field):
    try:
        result = Decimal(value.strip())
        if not result.is_finite():
            raise InvalidOperation
        return result
    except (InvalidOperation, AttributeError):
        raise ValueError(f"Riga {line}: {field} non è un numero finito: {value!r}") from None


def mm_text(value):
    if value == 0:
        return "0mm"
    result = format(value, "f")
    if "." in result:
        result = result.rstrip("0").rstrip(".")
    return result + "mm"


def read_positions(path, units="mm"):
    factor = Decimal("25.4") if units == "in" else Decimal(1)
    rows, seen = [], set()
    with Path(path).open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream, strict=True)
        if not reader.fieldnames:
            raise ValueError("CSV vuoto o privo di intestazioni.")
        reader.fieldnames = [h.strip() for h in reader.fieldnames]
        if len(set(reader.fieldnames)) != len(reader.fieldnames):
            raise ValueError("Intestazioni CSV duplicate.")
        missing = set(REQUIRED) - set(reader.fieldnames)
        if missing:
            raise ValueError("Colonne mancanti: " + ", ".join(sorted(missing)))
        for row in reader:
            line = reader.line_num
            if None in row or any(v is None for v in row.values()):
                raise ValueError(f"Riga {line}: numero di colonne errato.")
            if not any(v.strip() for v in row.values()):
                continue
            ref = row["Ref"].strip()
            if not ref:
                raise ValueError(f"Riga {line}: Ref vuoto.")
            if ref.casefold() in seen:
                raise ValueError(f"Riga {line}: riferimento duplicato {ref!r}.")
            seen.add(ref.casefold())
            side = row["Side"].strip().lower()
            if side not in ("top", "bottom"):
                raise ValueError(f"Riga {line}: Side deve essere top o bottom, ricevuto {side!r}.")
            x = number(row["PosX"], line, "PosX") * factor
            y = number(row["PosY"], line, "PosY") * factor
            rotation = number(row["Rot"], line, "Rot") % 360
            if rotation < 0:
                rotation += 360
            rows.append((ref, mm_text(x), mm_text(y), side.title(), float(rotation)))
    if not rows:
        raise ValueError("Il CSV non contiene componenti.")
    return rows


def write_workbook(rows, output, force=False):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError:
        raise ValueError("Manca openpyxl. Installalo con: python3 -m pip install openpyxl") from None
    output = Path(output)
    if output.suffix.lower() != ".xlsx":
        raise ValueError("Il file di output deve avere estensione .xlsx.")
    if output.exists() and not force:
        raise ValueError(f"Il file esiste già: {output}. Usa --force per sostituirlo.")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(HEADERS)
    for row in rows:
        sheet.append(row)
        # Forza testo letterale: un riferimento non deve diventare una formula.
        for cell in sheet[sheet.max_row][:4]:
            cell.data_type = "s"
    for cell in sheet[1]:
        cell.font = Font(name="Arial", size=11, bold=True)
    for column, width in zip("ABCDE", (18, 22, 22, 14, 14)):
        sheet.column_dimensions[column].width = width
    for cells in sheet.iter_rows(min_row=2, min_col=5, max_col=5):
        cells[0].number_format = "General"
    # Un errore di scrittura non deve lasciare un XLSX parziale.
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as tmp:
            temp_path = Path(tmp.name)
        workbook.save(temp_path)
        if force:
            os.replace(temp_path, output)
        else:
            # Pubblicazione senza sovrascrivere nemmeno in caso di creazione concorrente.
            os.link(temp_path, output)
    finally:
        workbook.close()
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", type=Path, help="CSV posizioni esportato da KiCad")
    parser.add_argument("-o", "--output", type=Path, help="Destinazione (default: <input>-JLCPCB.xlsx)")
    parser.add_argument("--units", choices=("mm", "in"), default="mm", help="Unità del CSV KiCad, default mm")
    parser.add_argument("--force", action="store_true", help="Sostituisci un XLSX già esistente")
    args = parser.parse_args(argv)
    output = args.output or args.input.with_name(args.input.stem + "-JLCPCB.xlsx")
    try:
        if output.resolve() == args.input.resolve():
            raise ValueError("Input e output devono essere file diversi.")
        rows = read_positions(args.input, args.units)
        write_workbook(rows, output, args.force)
    except (OSError, ValueError, csv.Error, UnicodeError, InvalidOperation) as exc:
        print(f"Errore: {exc}", file=sys.stderr)
        return 1
    print(f"Creato: {output.resolve()} ({len(rows)} componenti)")
    print("Coordinate preservate; rotazioni normalizzate in [0, 360). Verifica orientamenti nell'anteprima JLCPCB.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
