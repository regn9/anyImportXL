# anyImportXL Add-in (VBA)

This repository contains VBA source modules for an Excel add-in (`.xlam`) that imports HTML-disguised `.xls` reports, maps source metrics into variables (`A..Z`), evaluates output formulas (`O1..O10+`), and writes results into mapped target cells.

## Included modules

- `modEntry.bas` - Entry macro to open main UI.
- `modParserHtml.bas` - HTML table parser (`contentTable` fallback first table).
- `modModel.bas` - Model factory helpers for class-based records.
- `modFormula.bas` - Safe formula tokenizer + evaluator (+ - * / parentheses, A..Z).
- `modConfig.bas` - Persistent profile storage in add-in `zConfig` sheet (`VeryHidden`).
- `modWriter.bas` - Apply output values + `_ImportLog` logging.
- `CReportRow.cls`, `CVarBinding.cls`, `COutputDef.cls`, `CTargetMap.cls` - Class-based models used in collections.
- `frmMain.frm` - Main UserForm UI logic (VBE-importable form definition).

## Build `.xlam`

1. Open Excel and create a blank workbook.
2. Open VBA editor (`ALT+F11`).
3. Import modules (`File -> Import File...`):
   - `modEntry.bas`, `modParserHtml.bas`, `modModel.bas`, `modFormula.bas`, `modConfig.bas`, `modWriter.bas`.
   - `CReportRow.cls`, `CVarBinding.cls`, `COutputDef.cls`, `CTargetMap.cls`.
4. Import `frmMain.frm`.
5. Save workbook as **Excel Add-In (`.xlam`)**.
6. Install add-in:
   - Excel -> `File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go...`
   - Browse to your `.xlam`, check it, click OK.

## Run

- Run macro: `OpenAnyImportXlUI`.

## Manual test procedure

1. Open a target workbook with multiple sheets.
2. Launch add-in UI (`OpenAnyImportXlUI`).
3. Click **Import report...** and select your HTML `.xls` report.
4. Search/select a leaf row.
5. Assign variables (e.g., `A=Current`, `B=Prev`, `C=Change`).
6. Configure output rows:
   - `O1` formula: `A+B`
   - `O2` formula: `A*B/C`
   - Set target sheet + cell for each output.
7. Click **Preview** and verify computed values + `_ImportLog` warnings/errors.
8. Click **Apply** and verify mapped cells are overwritten by `.Value2` only.
9. Click **Save profile**, close/reopen UI, then **Load profile** and verify settings restore.

## Notes

- Per-workbook profile key is `ActiveWorkbook.FullName`.
- `zConfig` is created in the add-in workbook (`ThisWorkbook`) and set `xlSheetVeryHidden`.
- Section/header rows with blank numeric columns are ignored during import.

## UserForm export/import note (.frm + .frx)

When you export `frmMain` from VBE, Excel may generate both:
- `frmMain.frm` (text definition), and
- `frmMain.frx` (binary control/resource data, referenced by `OleObjectBlob`).

If `OleObjectBlob` is present in `.frm`, keep and distribute the matching `.frx` file beside it during import. Missing `.frx` can cause import failure or incomplete controls.

