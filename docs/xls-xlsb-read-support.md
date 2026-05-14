# XLS / XLSB complete read support bundle

This branch adds zero-dependency readers for legacy BIFF `.xls` workbooks and BIFF12 `.xlsb` workbooks, plus a rendered HTML smoke-check for the uploaded XLS fixture.

## Added public entry points

- `readXls(input, options?)` from `hucre/xls`
- `readXlsb(input, options?)` from `hucre/xlsb`
- high-level `read(input, options?)` auto-detects OLE2/CFB `.xls`, raw BIFF streams, OPC `.xlsb`, OPC `.xlsx`, and ODS.

## Implemented XLS coverage

- CFB/OLE2 FAT, DIFAT, MiniFAT, directory, root mini-stream extraction, and case-insensitive stream lookup.
- BIFF workbook globals, sheet metadata, SST with `Continue` records, code pages, 1900/1904 date system, custom formats, and XF number-format mapping.
- Correct BIFF8 compressed-Unicode handling: `fHighByte=0` decodes as low-byte UTF-16 code units, not as the workbook `CodePage`; this prevents ASCII/date strings from corrupting when `CodePage=1200`.
- BIFF5/BIFF8 sheet-name and label string handling.
- Numeric, RK, MulRK, string, bool/error, blank, formula cached values, formula token decompilation, merged cells, row metadata, and column metadata.
- Defined names / built-in names, local extern-sheet references, and basic HLINK hyperlink extraction.
- OLE Property Set document metadata from `\u0005SummaryInformation` and `\u0005DocumentSummaryInformation` streams.
- Password-protected XLS detection through FilePass and encrypted-package CFB streams.

## Implemented XLSB coverage

- OPC package and relationship traversal.
- Binary workbook sheet metadata and 1904 date-system flag.
- Binary shared strings and worksheet cell values.
- Binary formula cached values with formula-token decompilation from the formula payload tail.
- Binary styles parsing for custom number formats and CellXfs, including date conversion.
- Merged cells and basic binary hyperlink extraction.
- Password-protected XLSB CFB envelope detection.

## Uploaded XLS validation

The fixture `be2cfc00-79c3-469e-88b4-d979bd73805d.xls` was parsed with the new `readXls` implementation and rendered into:

- `examples/rendered/be2cfc00-79c3-469e-88b4-d979bd73805d.html`
- `artifacts/be2cfc00-79c3-469e-88b4-d979bd73805d-render.html`
- `artifacts/be2cfc00-79c3-469e-88b4-d979bd73805d.summary.json`
- `artifacts/be2cfc00-79c3-469e-88b4-d979bd73805d-check.md`

Observed sheet summary:

- `单选题`: 11 rows × 6 columns
- `多选题`: 14 rows × 15 columns
- `判断题`: 14 rows × 10 columns

A LibreOffice HTML export of the same file opened successfully and reported the same three sheet names. The hucre-rendered HTML preserves the sheet order and correctly renders mixed Chinese + ASCII strings such as `*正确答案`, the first answer value `B`, and date text `2022-04-01 00:00:00`.

## Validation status

- TypeScript syntax sanity check was run against the changed reader files in isolation.
- The uploaded XLS fixture was parsed end-to-end and rendered to HTML.
- `git bundle verify` was run on the produced bundle.

Apply the patch to the full hucre repository, then run the project-native checks:

```sh
pnpm lint
pnpm typecheck
pnpm test
pnpm build
```
