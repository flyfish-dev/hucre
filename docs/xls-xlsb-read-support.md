# XLS / XLSB complete read support bundle

This branch adds zero-dependency readers for legacy BIFF `.xls` workbooks and BIFF12 `.xlsb` workbooks.

## Added public entry points

- `readXls(input, options?)` from `hucre/xls`
- `readXlsb(input, options?)` from `hucre/xlsb`
- high-level `read(input, options?)` auto-detects OLE2/CFB `.xls`, raw BIFF streams, OPC `.xlsb`, OPC `.xlsx`, and ODS.

## Implemented XLS coverage

- CFB/OLE2 FAT, DIFAT, MiniFAT, directory, and mini-stream extraction.
- BIFF workbook globals, sheet metadata, SST with Continue records, codepages, date system, formats, XFs.
- BIFF5/BIFF8 sheet-name and label string handling.
- Numeric, RK, MulRK, string, bool/error, blank, formula cached values, formula token decompilation, merges, row/column metadata.
- Defined names / built-in names, local extern-sheet references, and basic HLINK hyperlink extraction.
- Password-protected XLS detection through FilePass.

## Implemented XLSB coverage

- OPC package and relationship traversal.
- Binary workbook sheet metadata and 1904 date-system flag.
- Binary shared strings and worksheet cell values.
- Binary formula cached values with formula-token decompilation from the formula payload tail.
- Binary styles parsing for custom number formats and CellXfs, including date conversion.
- Merged cells and basic binary hyperlink extraction.
- Password-protected XLSB CFB envelope detection.

## Validation status

The source files in this bundle were syntax-checked with TypeScript against the changed files in isolation. A full `pnpm test` / `pnpm typecheck` requires applying the bundle to the full hucre repository because the local environment could not clone GitHub.
