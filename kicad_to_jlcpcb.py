#!/usr/bin/env python3
"""Convert a KiCad position CSV to the JLCPCB CPL template's XLSX format.

Dependency: python3 -m pip install openpyxl
Usage: python3 kicad_to_jlcpcb.py board-all-pos.csv [-o board-CPL.xlsx]

KiCad coordinates are assumed to be in mm unless --units in is specified.
No axes are mirrored, the origin is not shifted, and no footprint corrections are applied.
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
        raise ValueError(f"Line {line}: {field} is not a finite number: {value!r}") from None


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
            raise ValueError("CSV is empty or has no headers.")
        reader.fieldnames = [h.strip() for h in reader.fieldnames]
        if len(set(reader.fieldnames)) != len(reader.fieldnames):
            raise ValueError("Duplicate CSV headers.")
        missing = set(REQUIRED) - set(reader.fieldnames)
        if missing:
            raise ValueError("Missing columns: " + ", ".join(sorted(missing)))
        for row in reader:
            line = reader.line_num
            if None in row or any(v is None for v in row.values()):
                raise ValueError(f"Line {line}: incorrect number of columns.")
            if not any(v.strip() for v in row.values()):
                continue
            ref = row["Ref"].strip()
            if not ref:
                raise ValueError(f"Line {line}: Ref is empty.")
            if ref.casefold() in seen:
                raise ValueError(f"Line {line}: duplicate reference {ref!r}.")
            seen.add(ref.casefold())
            side = row["Side"].strip().lower()
            if side not in ("top", "bottom"):
                raise ValueError(f"Line {line}: Side must be top or bottom, got {side!r}.")
            x = number(row["PosX"], line, "PosX") * factor
            y = number(row["PosY"], line, "PosY") * factor
            rotation = number(row["Rot"], line, "Rot") % 360
            if rotation < 0:
                rotation += 360
            rows.append((ref, mm_text(x), mm_text(y), side.title(), float(rotation)))
    if not rows:
        raise ValueError("CSV contains no components.")
    return rows


def write_workbook(rows, output, force=False):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError:
        raise ValueError("openpyxl is missing. Install it with: python3 -m pip install openpyxl") from None
    output = Path(output)
    if output.suffix.lower() != ".xlsx":
        raise ValueError("The output file must have an .xlsx extension.")
    if output.exists() and not force:
        raise ValueError(f"File already exists: {output}. Use --force to replace it.")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(HEADERS)
    for row in rows:
        sheet.append(row)
        # Force literal text: a reference must not become a formula.
        for cell in sheet[sheet.max_row][:4]:
            cell.data_type = "s"
    for cell in sheet[1]:
        cell.font = Font(name="Arial", size=11, bold=True)
    for column, width in zip("ABCDE", (18, 22, 22, 14, 14)):
        sheet.column_dimensions[column].width = width
    for cells in sheet.iter_rows(min_row=2, min_col=5, max_col=5):
        cells[0].number_format = "General"
    # A write error must not leave a partial XLSX file.
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as tmp:
            temp_path = Path(tmp.name)
        workbook.save(temp_path)
        if force:
            os.replace(temp_path, output)
        else:
            # Publish without overwriting, even if another process creates the file concurrently.
            os.link(temp_path, output)
    finally:
        workbook.close()
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", type=Path, help="Position CSV exported from KiCad")
    parser.add_argument("-o", "--output", type=Path, help="Output path (default: <input>-JLCPCB.xlsx)")
    parser.add_argument("--units", choices=("mm", "in"), default="mm", help="KiCad CSV units, default mm")
    parser.add_argument("--force", action="store_true", help="Replace an existing XLSX file")
    args = parser.parse_args(argv)
    output = args.output or args.input.with_name(args.input.stem + "-JLCPCB.xlsx")
    try:
        if output.resolve() == args.input.resolve():
            raise ValueError("Input and output must be different files.")
        rows = read_positions(args.input, args.units)
        write_workbook(rows, output, args.force)
    except (OSError, ValueError, csv.Error, UnicodeError, InvalidOperation) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(f"Created: {output.resolve()} ({len(rows)} components)")
    print("Coordinates preserved; rotations normalized to [0, 360). Check orientations in the JLCPCB preview.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
