import { describe, expect, it } from "vitest"
import { writeCfb } from "../src/xlsx/crypto/cfb"
import { readXls } from "../src/xls/reader"
import { read } from "../src/defter"

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
  EOF: 0x000a,
  FONT: 0x0031,
  DEFCOLWIDTH: 0x0055,
  COLINFO: 0x007d,
  DATEMODE: 0x0022,
  PALETTE: 0x0092,
  STANDARDWIDTH: 0x0099,
  BLANK: 0x0201,
  DIMENSIONS: 0x0200,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  ROW: 0x0208,
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

const bof = (dt: number): number[] =>
  record(SID.BOF, [...u16(0x0600), ...u16(dt), ...u16(0), ...u16(0), ...u32(0), ...u32(0)])
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
})
