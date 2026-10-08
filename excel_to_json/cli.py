import argparse
import sys
from pathlib import Path

from .converter import convert_workbook, write_json


def main(argv=None):
    p = argparse.ArgumentParser(description="Convert an Excel workbook to JSON.")
    p.add_argument("input", help=".xlsx file")
    p.add_argument("-o", "--output", help="output .json (default: next to input)")
    p.add_argument("-s", "--sheet", action="append", help="sheet to include (repeatable; default all)")
    p.add_argument("--header-row", type=int, default=1)
    p.add_argument("--drop-nulls", action="store_true", help="omit empty fields from each record")
    p.add_argument("--flat", action="store_true", help="single sheet: output a bare list, not {sheet: [...]}")
    a = p.parse_args(argv)

    data = convert_workbook(a.input, a.sheet, a.header_row, drop_null_fields=a.drop_nulls)
    if a.flat:
        if len(data) != 1:
            sys.exit("--flat needs exactly one sheet (use --sheet)")
        data = next(iter(data.values()))
    out = a.output or str(Path(a.input).with_suffix(".json"))
    write_json(data, out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
