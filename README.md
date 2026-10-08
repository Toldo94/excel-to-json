# Excel to JSON

A small desktop app (and CLI / Python library) that converts Excel workbooks to JSON.
Drop in an `.xlsx`, pick the sheets, preview the result, save.

## Download

Prebuilt apps for Windows, macOS and Linux are attached to each
[GitHub Release](../../releases). Builds are unsigned:

- **Windows:** SmartScreen shows a warning. Click *More info → Run anyway*.
- **macOS:** right-click the app and choose *Open* the first time. The build targets Apple Silicon.
- **Linux:** extract the `.tar.gz` and run `ExcelToJSON/ExcelToJSON`.

## Desktop app

1. Drag an `.xlsx` / `.xlsm` file onto the window, or click **Open…**.
2. Tick the sheets to include.
3. Adjust the **Header row** (default 1) and **Omit empty fields** if needed.
4. Check the live JSON preview, then **Copy** or **Save JSON…**.

## Command line

```sh
python -m excel_to_json.cli report.xlsx                 # all sheets -> report.json
python -m excel_to_json.cli report.xlsx -s Sheet1 --flat -o out.json
```

| Option | Meaning |
|---|---|
| `-o, --output` | Output path (default: next to the input, `.json`) |
| `-s, --sheet` | Sheet to include; repeatable (default: all) |
| `--header-row N` | Row containing column names (default 1) |
| `--drop-nulls` | Omit empty fields from each record |
| `--flat` | With a single sheet, output a bare list instead of `{sheet: [...]}` |

## Use from Python

```python
from excel_to_json.converter import convert_workbook

data = convert_workbook("report.xlsx")            # {"Sheet1": [{...}, ...], ...}
rows = convert_workbook("report.xlsx", sheets=["Sheet1"])["Sheet1"]
```

`convert_workbook(path, sheets=None, header_row=1, skip_empty_rows=True, drop_null_fields=False)`

## Output format

```json
{
  "People": [
    { "name": "Ana", "age": 30, "joined": "2024-01-05T00:00:00", "score": 1.5 },
    { "name": "Bo", "age": null, "joined": null, "score": 2 }
  ]
}
```

- Each sheet becomes a list of records keyed by the header row.
- Dates and times become ISO 8601 strings.
- Whole-number floats become integers (`30.0` → `30`).
- Empty cells become `null`; fully blank rows are skipped.
- Blank or duplicate headers become `column_N` or `name_2`.
- Formulas export their cached values. A file that was never opened and saved in Excel may show `null` for formulas.
- Only `.xlsx` / `.xlsm` are supported, not the old `.xls`.

## Development

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python main.py          # run the GUI
```

### Building

```sh
python build.py                   # -> dist/ExcelToJSON(.app | .exe)
```

PyInstaller can't cross-compile, so each platform has to be built on that platform.
The [build workflow](.github/workflows/build.yml) does this on GitHub Actions:
run it manually from the **Actions** tab, or push a tag (`git tag v0.1.0 && git push --tags`)
to build all three and publish a Release.

### Layout

```
excel_to_json/converter.py   core conversion (no GUI dependency)
excel_to_json/cli.py         command-line interface
excel_to_json/gui.py         PySide6 desktop app
main.py                      GUI entry point
build.py                     PyInstaller build script
```

**Stack:** Python 3.12, openpyxl, PySide6, PyInstaller.
