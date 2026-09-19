import { describe, expect, it } from "vitest"
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { readXls } from "../src/xls/reader"
import { read } from "../src/defter"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"

// ── Minimal BIFF8 .xls builder (test-only) ───────────────────────────

function concat(parts: Array<number[] | Uint8Array>): Uint8Array {
  let len = 0
  for (const p of parts) len += p.length
  const out = new Uint8Array(len)
  let off = 0
  for (const p of parts) {
    out.set(p instanceof Uint8Array ? p : new Uint8Array(p), off)
    off += p.length
  }
  return out
}
const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
const u32 = (n: number): number[] => [
  n & 0xff,
  (n >> 8) & 0xff,
  (n >> 16) & 0xff,
  (n >>> 24) & 0xff,
]
function f64(n: number): number[] {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return [...b]
}
// XLUnicodeString (u16 cch) / ShortXLUnicodeString (u8 cch), compressed
const xlStr = (s: string): number[] => [...u16(s.length), 0, ...[...s].map((c) => c.charCodeAt(0))]
const shortStr = (s: string): number[] => [s.length, 0, ...[...s].map((c) => c.charCodeAt(0))]
function record(sid: number, data: number[]): number[] {
  return [...u16(sid), ...u16(data.length), ...data]
}
const rkInt = (v: number): number[] => u32(((v << 2) | 2) >>> 0)

const SID = {
  FORMULA: 0x0006,
  STRING: 0x0207,
  EOF: 0x000a,
  HORIZONTALPAGEBREAKS: 0x001b,
  VERTICALPAGEBREAKS: 0x001a,
  LEFTMARGIN: 0x0026,
  RIGHTMARGIN: 0x0027,
  TOPMARGIN: 0x0028,
  BOTTOMMARGIN: 0x0029,
  PRINTHEADERS: 0x002a,
  PRINTGRIDLINES: 0x002b,
  FONT: 0x0031,
  DEFCOLWIDTH: 0x0055,
  COLINFO: 0x007d,
  WSBOOL: 0x0081,
  HCENTER: 0x0083,
  VCENTER: 0x0084,
  DATEMODE: 0x0022,
  PALETTE: 0x0092,
  STANDARDWIDTH: 0x0099,
  SCL: 0x00a0,
  SETUP: 0x00a1,
  BLANK: 0x0201,
  MULBLANK: 0x00be,
  DIMENSIONS: 0x0200,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  ROW: 0x0208,
  WINDOW2: 0x023e,
  RK: 0x027e,
  MULRK: 0x00bd,
  LABELSST: 0x00fd,
  SST: 0x00fc,
  XF: 0x00e0,
  BOUNDSHEET: 0x0085,
  MERGECELLS: 0x00e5,
  DEFAULTROWHEIGHT: 0x0225,
  BOF: 0x0809,
}

const bof = (dt: number, version = 0x0600): number[] =>
  record(SID.BOF, [...u16(version), ...u16(dt), ...u16(0), ...u16(0), ...u32(0), ...u32(0)])
const eof = (): number[] => record(SID.EOF, [])

function sstRecord(strings: string[]): number[] {
  const body: number[] = [...u32(strings.length), ...u32(strings.length)]
  for (const s of strings) body.push(...u16(s.length), 0, ...[...s].map((c) => c.charCodeAt(0)))
  return record(SID.SST, body)
}

function buildXls(opts: { dateFmtId?: number } = {}): Uint8Array {
  const strings = ["Name", "Score", "Ada"]

  const sheet = concat([
    bof(0x0010),
    record(SID.LABELSST, [...u16(0), ...u16(0), ...u16(0), ...u32(0)]),
    record(SID.LABELSST, [...u16(0), ...u16(1), ...u16(0), ...u32(1)]),
    record(SID.LABELSST, [...u16(1), ...u16(0), ...u16(0), ...u32(2)]),
    record(SID.RK, [...u16(1), ...u16(1), ...u16(0), ...rkInt(95)]),
    record(SID.NUMBER, [...u16(1), ...u16(2), ...u16(0), ...f64(3.14)]),
    record(SID.NUMBER, [...u16(1), ...u16(3), ...u16(1), ...f64(45000)]), // date xf
    record(SID.LABEL, [...u16(2), ...u16(0), ...u16(0), ...xlStr("Hi")]),
    record(SID.BOOLERR, [...u16(2), ...u16(1), ...u16(0), 1, 0]), // true
    record(SID.BOOLERR, [...u16(2), ...u16(2), ...u16(0), 0x07, 1]), // #DIV/0!
    record(SID.MULRK, [
      ...u16(3),
      ...u16(0),
      ...u16(0),
      ...rkInt(10),
      ...u16(0),
      ...rkInt(20),
      ...u16(1),
    ]),
    record(SID.MERGECELLS, [...u16(1), ...u16(0), ...u16(0), ...u16(0), ...u16(1)]),
    eof(),
  ])

  // Globals — BOUNDSHEET position is filled in once the globals size is known.
  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof(0x0005),
      record(SID.DATEMODE, u16(0)),
      record(SID.XF, [...u16(0), ...u16(0), ...Array.from({ length: 16 }, () => 0)]), // general
      // The date xf's built-in format id is a parameter: the built-in date
      // set is wider than the familiar 14-22 block. See the CJK case below.
      record(SID.XF, [
        ...u16(0),
        ...u16(opts.dateFmtId ?? 14),
        ...Array.from({ length: 16 }, () => 0),
      ]),
      sstRecord(strings),
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Sheet1")]),
      eof(),
    ])

  const globalsLen = makeGlobals(0).length
  const globals = makeGlobals(globalsLen)
  const workbookStream = concat([globals, sheet])

  return writeCfb([{ name: "Workbook", data: workbookStream }])
}

function buildSingleSheetXls(sheetRecords: number[][], globalRecords: number[][] = []): Uint8Array {
  const sheet = concat([bof(0x0010), ...sheetRecords, eof()])
  const globals = (sheetPos: number) =>
    concat([
      bof(0x0005),
      ...globalRecords,
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Cases")]),
      eof(),
    ])
  return writeCfb([{ name: "Workbook", data: concat([globals(globals(0).length), sheet]) }])
}

function buildStyledXls(): Uint8Array {
  const font = (name: string, heightTwips: number, color: number, weight: number): number[] =>
    record(SID.FONT, [
      ...u16(heightTwips),
      ...u16(0),
      ...u16(color),
      ...u16(weight),
      ...u16(0),
      0,
      2,
      0,
      0,
      ...shortStr(name),
    ])
  const xf = (
    fontIndex: number,
    alignment: number,
    border1: number,
    border2: number,
    pattern: number,
  ): number[] =>
    record(SID.XF, [
      ...u16(fontIndex),
      ...u16(0),
      ...u16(1),
      alignment,
      0,
      0,
      0x78,
      ...u32(border1),
      ...u32(border2),
      ...u16(pattern),
    ])

  const sheet = concat([
    bof(0x0010),
    record(SID.DEFAULTROWHEIGHT, [...u16(0), ...u16(285)]),
    record(SID.DEFCOLWIDTH, u16(8)),
    record(SID.STANDARDWIDTH, u16(9 * 256)),
    record(SID.COLINFO, [
      ...u16(0),
      ...u16(0),
      ...u16(20.5 * 256),
      ...u16(1),
      ...u16(0x0002),
      ...u16(0),
    ]),
    record(SID.DIMENSIONS, [...u32(0), ...u32(1), ...u16(0), ...u16(2), ...u16(0)]),
    record(SID.ROW, [...u16(0), ...u16(0), ...u16(2), ...u16(600), ...u32(0), ...u32(0x0040)]),
    record(SID.LABELSST, [...u16(0), ...u16(0), ...u16(1), ...u32(0)]),
    record(SID.BLANK, [...u16(0), ...u16(1), ...u16(1)]),
    eof(),
  ])

  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof(0x0005),
      font("Default", 220, 0x7fff, 400),
      font("Demo", 280, 8, 700),
      xf(0, 0x20, 0, 0, 0x20c0),
      // Center + vertical center + wrap, thin borders, solid custom palette fill.
      xf(1, 0x1a, 0x20401111, 0x04002040, 0x2008),
      record(SID.PALETTE, [...u16(1), 0x12, 0x34, 0x56, 0]),
      sstRecord(["Styled"]),
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Styled")]),
      eof(),
    ])

  const globalsLength = makeGlobals(0).length
  return writeCfb([{ name: "Workbook", data: concat([makeGlobals(globalsLength), sheet]) }])
}

function buildRichTextXls(): Uint8Array {
  const font = (name: string, color: number): number[] =>
    record(SID.FONT, [
      ...u16(220),
      ...u16(0),
      ...u16(color),
      ...u16(400),
      ...u16(0),
      0,
      2,
      0,
      0,
      ...shortStr(name),
    ])
  const text = "BlackRed"
  const richSst = record(SID.SST, [
    ...u32(1),
    ...u32(1),
    ...u16(text.length),
    0x08,
    ...u16(2),
    ...[...text].map((character) => character.charCodeAt(0)),
    ...u16(0),
    ...u16(0),
    ...u16(5),
    ...u16(1),
  ])
  const sheet = concat([
    bof(0x0010),
    record(SID.LABELSST, [...u16(0), ...u16(0), ...u16(0), ...u32(0)]),
    eof(),
  ])
  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof(0x0005),
      font("Default", 0x7fff),
      font("Accent", 8),
      record(SID.XF, [...u16(0), ...u16(0), ...Array.from({ length: 16 }, () => 0)]),
      record(SID.PALETTE, [...u16(1), 0xaa, 0x11, 0x22, 0]),
      richSst,
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Rich")]),
      eof(),
    ])

  const globalsLength = makeGlobals(0).length
  return writeCfb([{ name: "Workbook", data: concat([makeGlobals(globalsLength), sheet]) }])
}

function buildPageBreakPreviewXls(
  options: {
    compactLegacyBreaks?: boolean
    rowBreakIds?: number[]
    colBreakIds?: number[]
    includeScl?: boolean
    pageBreakPreview?: boolean
    pageBreakZoom?: number
    normalZoom?: number
    setupFlags?: number
    windowFlags?: number
    rowBreakData?: number[]
    colBreakData?: number[]
  } = {},
): Uint8Array {
  const {
    compactLegacyBreaks = false,
    rowBreakIds = [7],
    colBreakIds = [2],
    includeScl = true,
    pageBreakPreview = true,
    pageBreakZoom = 130,
    normalZoom = 90,
    setupFlags = 0x0002,
    windowFlags = pageBreakPreview ? 0x0eb6 : 0x06b6,
    rowBreakData,
    colBreakData,
  } = options
  const version = compactLegacyBreaks ? 0x0500 : 0x0600
  const breakData = (ids: number[], maximumSpan: number): number[] => [
    ...u16(ids.length),
    ...ids.flatMap((id) => [
      ...u16(id),
      ...(compactLegacyBreaks ? [] : [...u16(0), ...u16(maximumSpan)]),
    ]),
  ]
  const records = [
    bof(0x0010, version),
    // Page Break Preview, grid/headers visible, 60% zoom through SCL.
    record(SID.WINDOW2, [
      ...u16(windowFlags),
      ...u16(0),
      ...u16(0),
      ...u16(64),
      ...u16(0),
      ...u16(pageBreakZoom),
      ...u16(normalZoom),
      ...u16(0),
      ...u16(0),
    ]),
    record(SID.WSBOOL, u16(0x0100)),
    record(SID.LEFTMARGIN, f64(0.75)),
    record(SID.RIGHTMARGIN, f64(0.75)),
    record(SID.TOPMARGIN, f64(1)),
    record(SID.BOTTOMMARGIN, f64(1)),
    record(SID.PRINTHEADERS, u16(1)),
    record(SID.PRINTGRIDLINES, u16(1)),
    record(SID.HCENTER, u16(1)),
    record(SID.VCENTER, u16(0)),
    record(SID.SETUP, [
      ...u16(9),
      ...u16(100),
      ...u16(1),
      ...u16(1),
      ...u16(2),
      ...u16(setupFlags),
      ...u16(180),
      ...u16(180),
      ...f64(0.5),
      ...f64(0.5),
      ...u16(1),
    ]),
    record(SID.HORIZONTALPAGEBREAKS, rowBreakData ?? breakData(rowBreakIds, 255)),
    record(SID.VERTICALPAGEBREAKS, colBreakData ?? breakData(colBreakIds, 65_535)),
    record(SID.LABEL, [...u16(0), ...u16(0), ...u16(0), ...xlStr("Preview")]),
    eof(),
  ]
  if (includeScl) records.splice(2, 0, record(SID.SCL, [...u16(3), ...u16(5)]))
  const sheet = concat(records)
  const makeGlobals = (sheetPos: number): Uint8Array =>
    concat([
      bof(0x0005, version),
      record(SID.BOUNDSHEET, [...u32(sheetPos), 0, 0, ...shortStr("Preview")]),
      eof(),
    ])
  const globalsLength = makeGlobals(0).length
  return writeCfb([{ name: "Workbook", data: concat([makeGlobals(globalsLength), sheet]) }])
}

describe("XLS (BIFF8) reader", () => {
  it("decodes SST labels, RK, MULRK, numbers, bools, errors, dates, and merges", async () => {
    const wb = await readXls(buildXls())
    expect(wb.sheets.length).toBe(1)
    expect(wb.sheets[0].name).toBe("Sheet1")
    const rows = wb.sheets[0].rows
    // Padded to the sheet width, not to this row's own last cell:
    // `rows` is a dense rectangle, which these readers used to leave
    // ragged while readXlsx did not. See #494.
    expect(rows[0]).toEqual(["Name", "Score", null, null])
    expect(rows[1][0]).toBe("Ada")
    expect(rows[1][1]).toBe(95)
    expect(rows[1][2]).toBeCloseTo(3.14, 5)
    expect(rows[1][3]).toBeInstanceOf(Date)
    expect(rows[2][0]).toBe("Hi")
    expect(rows[2][1]).toBe(true)
    expect(rows[2][2]).toBe("#DIV/0!")
    expect(rows[3][0]).toBe(10)
    expect(rows[3][1]).toBe(20)
    expect(wb.sheets[0].merges).toEqual([{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }])
  })

  // ── #439: the built-in date set is wider than 14-22 / 45-47 ──────────
  // Built-ins 27-36 (CJK) and 50-58 (Thai/Chinese/Korean) are date and
  // time formats, and they carry no FORMAT record — so a reader that does
  // not know them falls through to "not a date" and hands back the raw
  // serial. This reader used to keep a 12-entry table of its own.
  describe("built-in date format ids outside the familiar block", () => {
    const CJK_AND_EXTENDED = [
      27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58,
    ]

    for (const id of CJK_AND_EXTENDED) {
      it(`reads a cell styled with built-in format ${id} as a Date`, async () => {
        const wb = await readXls(buildXls({ dateFmtId: id }))

        expect(wb.sheets[0].rows[1][3]).toBeInstanceOf(Date)
      })
    }

    it("still treats a non-date built-in as a number", async () => {
      const wb = await readXls(buildXls({ dateFmtId: 3 }))

      expect(wb.sheets[0].rows[1][3]).toBe(45000)
    })
  })

  it("is auto-detected by read()", async () => {
    const wb = await read(buildXls())
    expect(wb.sheets[0].rows[1][0]).toBe("Ada")
    expect(wb.sheets[0].rows[1][1]).toBe(95)
  })

  it("preserves BIFF column widths, row heights, style-only cells, and cell formatting", async () => {
    const workbook = await readXls(buildStyledXls(), { readStyles: true })
    const sheet = workbook.sheets[0]!

    expect(sheet.sheetFormat).toMatchObject({ defaultRowHeight: 14.25, defaultColWidth: 9 })
    expect(sheet.columns?.[0]).toMatchObject({ width: 20.5 })
    expect(sheet.rowDefs?.get(0)).toMatchObject({ height: 30, customHeight: true })
    expect(sheet.cells?.get("0,0")).toMatchObject({
      value: "Styled",
      type: "string",
      style: {
        font: { name: "Demo", size: 14, bold: true, color: { rgb: "123456" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "123456" } },
        border: {
          left: { style: "thin" },
          right: { style: "thin" },
          top: { style: "thin" },
          bottom: { style: "thin" },
        },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
      },
    })
    expect(sheet.cells?.get("0,1")).toMatchObject({ value: null, type: "empty" })
  })

  it("preserves BIFF8 SST rich-text font runs", async () => {
    const workbook = await readXls(buildRichTextXls(), { readStyles: true })

    expect(workbook.sheets[0].cells?.get("0,0")).toMatchObject({
      value: "BlackRed",
      type: "richText",
      richText: [
        { text: "Black", font: { name: "Default", size: 11 } },
        { text: "Red", font: { name: "Accent", size: 11, color: { rgb: "AA1122" } } },
      ],
    })
  })

  it("preserves BIFF page-break view, zoom, print geometry, and explicit breaks", async () => {
    const workbook = await readXls(buildPageBreakPreviewXls())
    const sheet = workbook.sheets[0]!

    expect(sheet.view).toEqual({ mode: "pageBreakPreview", zoomScale: 60 })
    expect(sheet.pageSetup).toEqual({
      paperSize: 9,
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 2,
      scale: 100,
      margins: { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
      showGridLines: true,
      showRowColHeaders: true,
      horizontalCentered: true,
      copies: 1,
      horizontalDpi: 180,
      verticalDpi: 180,
    })
    // Hucre exposes the zero-based row/column before the break, matching its
    // long-standing SpreadsheetML reader/writer contract.
    expect(sheet.rowBreaks).toEqual([6])
    expect(sheet.colBreaks).toEqual([1])
  })

  it("preserves compact BIFF5/7 page-break entries", async () => {
    const sheet = (await readXls(buildPageBreakPreviewXls({ compactLegacyBreaks: true })))
      .sheets[0]!
    expect(sheet.rowBreaks).toEqual([6])
    expect(sheet.colBreaks).toEqual([1])
  })

  it("uses the active WINDOW2 zoom when an associated SCL is absent", async () => {
    const preview = (
      await readXls(
        buildPageBreakPreviewXls({ includeScl: false, pageBreakZoom: 130, normalZoom: 90 }),
      )
    ).sheets[0]!
    const normal = (
      await readXls(
        buildPageBreakPreviewXls({
          includeScl: false,
          pageBreakPreview: false,
          pageBreakZoom: 130,
          normalZoom: 90,
        }),
      )
    ).sheets[0]!

    expect(preview.view).toMatchObject({ mode: "pageBreakPreview", zoomScale: 130 })
    expect(normal.view).toEqual({ zoomScale: 90 })
  })

  it("normalizes BIFF break boundaries, duplicates, and invalid positions", async () => {
    const sheet = (
      await readXls(
        buildPageBreakPreviewXls({
          rowBreakIds: [0, 1, 1, 65_535],
          colBreakIds: [0, 1, 1, 255, 256],
        }),
      )
    ).sheets[0]!

    // Raw BIFF values name the first item after a break. Zero cannot name a
    // preceding item; the maximum values put the break immediately before
    // the final BIFF8 row/column.
    expect(sheet.rowBreaks).toEqual([0, 65_534])
    expect(sheet.colBreaks).toEqual([0, 254])
  })

  it("keeps empty BIFF break collections absent", async () => {
    const sheet = (await readXls(buildPageBreakPreviewXls({ rowBreakIds: [], colBreakIds: [] })))
      .sheets[0]!
    expect(sheet.rowBreaks).toBeUndefined()
    expect(sheet.colBreaks).toBeUndefined()
  })

  it("rejects truncated or overlong BIFF8 horizontal and vertical break records", async () => {
    const completeHorizontal = [...u16(1), ...u16(7), ...u16(0), ...u16(255)]
    const completeVertical = [...u16(1), ...u16(2), ...u16(0), ...u16(65_535)]

    await expect(
      readXls(
        buildPageBreakPreviewXls({
          // This has the exact shape of a compact BIFF5/7 entry. A BIFF8
          // reader must not reinterpret it after the span bytes were lost.
          rowBreakData: [...u16(1), ...u16(7)],
        }),
      ),
    ).rejects.toThrow(/horizontal page-break record: expected 8 bytes, got 4/)
    await expect(
      readXls(
        buildPageBreakPreviewXls({
          rowBreakData: [...completeHorizontal, 0],
        }),
      ),
    ).rejects.toThrow(/horizontal page-break record: expected 8 bytes, got 9/)
    await expect(
      readXls(
        buildPageBreakPreviewXls({
          colBreakData: [...u16(1), ...u16(2)],
        }),
      ),
    ).rejects.toThrow(/vertical page-break record: expected 8 bytes, got 4/)
    await expect(
      readXls(
        buildPageBreakPreviewXls({
          colBreakData: [...completeVertical, 0],
        }),
      ),
    ).rejects.toThrow(/vertical page-break record: expected 8 bytes, got 9/)
  })

  it("keeps XLS break semantics when converted through an XLSX round-trip", async () => {
    const xlsSheet = (
      await readXls(buildPageBreakPreviewXls({ rowBreakIds: [1, 7], colBreakIds: [1, 2] }))
    ).sheets[0]!
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: xlsSheet.name,
          rows: xlsSheet.rows,
          rowBreaks: xlsSheet.rowBreaks,
          colBreaks: xlsSheet.colBreaks,
        },
      ],
    })
    const roundTripped = (await readXlsx(xlsx)).sheets[0]!

    expect(xlsSheet.rowBreaks).toEqual([0, 6])
    expect(xlsSheet.colBreaks).toEqual([0, 1])
    expect(roundTripped.rowBreaks).toEqual(xlsSheet.rowBreaks)
    expect(roundTripped.colBreaks).toEqual(xlsSheet.colBreaks)
  })

  it("accepts BIFF5/7 (0x0500) workbooks through the extended reader", async () => {
    // A globals BOF declaring BIFF version 0x0500 (Excel 5/95), then EOF.
    const biff5Bof = record(SID.BOF, [
      ...u16(0x0500),
      ...u16(0x0005),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(0),
    ])
    const stream = concat([biff5Bof, eof()])
    const data = writeCfb([{ name: "Workbook", data: stream }])
    const workbook = await readXls(data)
    expect(workbook.sheets).toEqual([{ name: "Sheet1", rows: [] }])
  })

  it("retains the full BIFF print setup flag semantics", async () => {
    const detailed = (
      await readXls(
        buildPageBreakPreviewXls({
          setupFlags: 0x0001 | 0x0002 | 0x0008 | 0x0010 | 0x0020 | 0x0080 | 0x0200 | 0x0400,
          includeScl: false,
        }),
      )
    ).sheets[0]!.pageSetup!
    expect(detailed).toMatchObject({
      pageOrder: "overThenDown",
      orientation: "portrait",
      blackAndWhite: true,
      draft: true,
      cellComments: "atEnd",
      errors: "blank",
      firstPageNumber: 1,
      useFirstPageNumber: true,
    })
    const noPrinter = (
      await readXls(
        buildPageBreakPreviewXls({
          setupFlags: 0x0004 | 0x0020 | 0x0800,
        }),
      )
    ).sheets[0]!.pageSetup!
    expect(noPrinter).toMatchObject({ cellComments: "asDisplayed", errors: "dash" })
    expect(noPrinter.paperSize).toBeUndefined()
    expect(noPrinter.orientation).toBeUndefined()
    const noOrientation = (
      await readXls(
        buildPageBreakPreviewXls({
          setupFlags: 0x0040 | 0x0c00,
        }),
      )
    ).sheets[0]!.pageSetup!
    expect(noOrientation.errors).toBe("NA")
    expect(noOrientation.orientation).toBeUndefined()
  })

  it("uses window flags for right-to-left and hidden headers without treating normal view as a page preview", async () => {
    const sheet = (
      await readXls(
        buildPageBreakPreviewXls({
          pageBreakPreview: false,
          windowFlags: 0x0040,
          includeScl: false,
        }),
      )
    ).sheets[0]!
    expect(sheet.view).toMatchObject({
      rightToLeft: true,
      showGridLines: false,
      showRowColHeaders: false,
      zoomScale: 90,
    })
    expect(sheet.view?.mode).toBeUndefined()
  })

  it("decodes numeric, boolean, error, string and blank formula caches", async () => {
    const formula = (row: number, cached: number[]) =>
      record(SID.FORMULA, [...u16(row), ...u16(0), ...u16(0), ...cached])
    const special = (kind: number, value: number) => [kind, 0, value, 0, 0, 0, 0xff, 0xff]
    const data = buildSingleSheetXls([
      formula(0, f64(3.5)),
      formula(1, special(1, 1)),
      formula(2, special(2, 0x07)),
      formula(3, special(0, 0)),
      record(SID.STRING, xlStr("cached")),
      formula(4, special(3, 0)),
      record(SID.BOOLERR, [...u16(5), ...u16(0), ...u16(0), 0xff, 1]),
    ])
    const valueOnly = (await readXls(data)).sheets[0]!
    expect(valueOnly.rows.slice(0, 6).map((row) => row[0])).toEqual([
      3.5,
      true,
      "#DIV/0!",
      "cached",
      null,
      "#ERR!",
    ])
    const styled = (await readXls(data, { readStyles: true })).sheets[0]!
    expect(styled.cells?.get("4,0")?.type).toBe("empty")
  })

  it("ignores blank records outside declared dimensions while preserving valid null cells", async () => {
    const sheet = (
      await readXls(
        buildSingleSheetXls(
          [
            record(SID.DIMENSIONS, [...u32(0), ...u32(1), ...u16(0), ...u16(1), ...u16(0)]),
            record(SID.BLANK, [...u16(0), ...u16(0), ...u16(0)]),
            record(SID.BLANK, [...u16(0), ...u16(1), ...u16(0)]),
            record(SID.MULBLANK, [...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u16(1)]),
          ],
          [record(SID.PALETTE, [0])],
        ),
        { readStyles: true },
      )
    ).sheets[0]!
    expect(sheet.rows[0]).toEqual([null])
    expect(sheet.cells?.has("0,1")).toBe(false)
  })

  it("decodes BIFF font, alignment, border and pattern variants from independent XFs", async () => {
    const font = (name: string, underline: number, escapement: number, flags = 0x000a) =>
      record(SID.FONT, [
        ...u16(240),
        ...u16(flags),
        ...u16(8),
        ...u16(700),
        ...u16(escapement),
        underline,
        2,
        1,
        0,
        ...shortStr(name),
      ])
    const xf = (
      fontIndex: number,
      fmt: number,
      protection: number,
      alignment: number,
      rotation: number,
      indent: number,
      border1: number,
      border2: number,
      colors: number,
    ) =>
      record(SID.XF, [
        ...u16(fontIndex),
        ...u16(fmt),
        ...u16(protection),
        alignment,
        rotation,
        indent,
        0,
        ...u32(border1),
        ...u32(border2),
        ...u16(colors),
      ])
    const palette = record(SID.PALETTE, [...u16(2), 0x10, 0x20, 0x30, 0, 0x40, 0x50, 0x60, 0])
    const sheetRecords = Array.from({ length: 6 }, (_, row) =>
      record(SID.LABEL, [...u16(row), ...u16(0), ...u16(row), ...xlStr("x")]),
    )
    const bytes = buildSingleSheetXls(sheetRecords, [
      font("Default", 0, 0, 0),
      font("Double", 2, 1),
      font("Accounting", 0x21, 2),
      font("DoubleAccounting", 0x22, 0),
      font("Skipped4", 1, 0),
      record(SID.FONT, [0]),
      xf(0, 0, 1, 0, 0, 0, 0, 0, 0),
      xf(1, 0, 2, 0x1a, 45, 0x51, 0xc0084321, (1 << 26) | (5 << 21) | (8 << 14) | 8, 8 | (9 << 7)),
      xf(2, 0, 1, 0, 0, 0x80, 0, 63 << 26, 0),
      xf(3, 14, 1, 0, 0, 0, 0, 0, 0),
      xf(5, 0, 1, 0, 0, 0, 0, 0, 0),
      xf(0, 0, 1, 0, 0, 0, 0, 1 << 26, 0x7f | (0x7f << 7)),
      palette,
    ])
    const cells = (await readXls(bytes, { readStyles: true })).sheets[0]!.cells!
    const detailed = cells.get("1,0")!.style!
    expect(detailed.font).toMatchObject({
      name: "Double",
      bold: true,
      italic: true,
      strikethrough: true,
      underline: "double",
      vertAlign: "superscript",
      family: 2,
      charset: 1,
      color: { rgb: "102030" },
    })
    expect(detailed.alignment).toMatchObject({
      horizontal: "center",
      vertical: "center",
      wrapText: true,
      shrinkToFit: true,
      indent: 1,
      textRotation: 45,
      readingOrder: "ltr",
    })
    expect(detailed.border?.left?.style).toBe("thin")
    expect(detailed.border?.diagonalUp).toBe(true)
    expect(detailed.border?.diagonalDown).toBe(true)
    expect(detailed.fill).toMatchObject({
      pattern: "solid",
      fgColor: { rgb: "102030" },
      bgColor: { rgb: "405060" },
    })
    expect(cells.get("2,0")!.style?.font).toMatchObject({
      underline: "singleAccounting",
      vertAlign: "subscript",
    })
    expect(cells.get("2,0")!.style?.alignment?.readingOrder).toBe("rtl")
    expect(cells.get("2,0")!.style?.fill).toBeUndefined()
    expect(cells.get("3,0")!.style?.font?.underline).toBe("doubleAccounting")
    expect(cells.get("3,0")!.style?.numFmt).toBeTruthy()
    expect(cells.get("4,0")!.style?.font?.name).toBe("Skipped4")
    expect(cells.get("5,0")!.style?.fill).toMatchObject({ pattern: "solid" })
    const noPaletteFill = cells.get("5,0")!.style?.fill
    expect(noPaletteFill?.type).toBe("pattern")
    if (noPaletteFill?.type === "pattern") expect(noPaletteFill.fgColor).toBeUndefined()
  })
})
