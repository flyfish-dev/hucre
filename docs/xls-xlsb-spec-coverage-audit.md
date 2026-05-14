# XLS / XLSB spec coverage audit and fixture render

This branch extends the previous XLS / XLSB reader work with another pass over the Microsoft binary workbook model and adds a real `.xls` fixture render output for usability checks.

## Additional implementation work in this pass

### XLS / BIFF + CFB

- Added parsing for OLE property-set streams in CFB containers:
  - `\u0005SummaryInformation`
  - `\u0005DocumentSummaryInformation`
- Mapped parsed metadata into `Workbook.properties`:
  - title, subject, creator, keywords, description, lastModifiedBy
  - created, modified
  - category, manager, company
- Kept XLS value parsing zero-dependency and routed property parsing through `readXls`, so normal `read()` callers receive the metadata as well.

### XLSB / BIFF12 + OPC

- Added `docProps/core.xml`, `docProps/app.xml`, and `docProps/custom.xml` parsing to `readXlsb`, using the same property readers already used by the XLSX reader.
- XLSB now returns `Workbook.properties` in addition to sheets/date system/cell values.

## Uploaded XLS render verification

The uploaded file `be2cfc00-79c3-469e-88b4-d979bd73805d.xls` was parsed using the XLS reader and rendered to:

- `examples/rendered/be2cfc00-79c3-469e-88b4-d979bd73805d.html`

Observed workbook structure:

| Sheet | Rows | Columns | Status |
| --- | ---: | ---: | --- |
| 单选题 | 11 | 6 | Parsed and rendered |
| 多选题 | 14 | 15 | Parsed and rendered |
| 判断题 | 14 | 10 | Parsed and rendered |

Parsed CFB document properties from the same workbook:

| Property | Value |
| --- | --- |
| creator | DingTalk |
| lastModifiedBy | DingTalk |
| created | 2006-09-16T00:00:00.000Z |
| modified | 2025-10-21T15:23:14.000Z |

The HTML render preserves Chinese text, line breaks, blank cells, correct-answer columns, and sheet separation. This gives a concrete smoke test for the CFB -> BIFF -> Workbook -> HTML path.

## Remaining guardrails

- Password-protected/encrypted XLS/XLSB files are detected and surfaced as `EncryptedFileError`; decryption is intentionally not performed by the zero-dependency reader.
- Chart/drawing/pivot/comment authoring is outside the core read-value path. Existing XLSX round-trip preservation remains separate from these XLS/XLSB binary readers.
