# XLS/XLSB parser render check

This artifact set was produced from the uploaded workbook `be2cfc00-79c3-469e-88b4-d979bd73805d.xls` using the zero-dependency XLS reader path (`readXls`) after the BIFF/CFB fixes in this bundle.

## Parsed workbook summary

| Sheet | Rows | Columns | Merge ranges |
| --- | ---: | ---: | ---: |
| 单选题 | 11 | 6 | 0 |
| 多选题 | 14 | 15 | 1 |
| 判断题 | 14 | 10 | 1 |

The generated HTML is stored at:

- `examples/rendered/be2cfc00-79c3-469e-88b4-d979bd73805d.html`

## Usability checks covered

- OLE2/CFB stream extraction: `Workbook`, `\u0005SummaryInformation`, and `\u0005DocumentSummaryInformation` streams are discoverable.
- BIFF8 workbook globals: CodePage `1200`, BoundSheet metadata, SST/Continue strings, row/column records, and merge ranges.
- BIFF8 compressed Unicode: compressed XLUnicodeString payloads are decoded as low-byte UTF-16 code units instead of through the workbook CodePage. This fixes ASCII/date-like text when CodePage is `1200`.
- HTML rendering: all three sheets render with Chinese text, multi-line strings, and merged instruction cells.
