// ── XLS (BIFF8) Reader ───────────────────────────────────────────────
// Read legacy Excel 97-2003 .xls files: an OLE2/CFB container whose
// "Workbook" stream is a BIFF8 record sequence. Reuses the CFB reader
// (shared with encryption) and decodes the records into the standard
// Workbook model. Read-only (MS-XLS).

import type {
  AlignmentStyle,
  BorderLineStyle,
  BorderSide,
  Cell,
  CellStyle,
  CellType,
  CellValue,
  Color,
  FillPattern,
  FontStyle,
  MergeRange,
  ReadOptions,
  RichTextRun,
  RowDef,
  Sheet,
  Workbook,
} from "../_types"
import { ParseError } from "../errors"
import { MAX_COL_INDEX, MAX_ROW_INDEX, MAX_TOTAL_CELLS } from "../limits"
import { readInputToUint8Array } from "../_input"
import { readCfb } from "../xlsx/crypto/cfb"
import { isBuiltinDateFormatId, isDateFormat, serialToDate } from "../_date"
import { DEFAULT_INDEXED_PALETTE } from "../xlsx/indexed-palette"
import { BUILTIN_NUM_FMTS } from "../xlsx/styles"
import {
  decodeRk,
  parseRecords,
  parseSstEntries,
  Reader,
  SID,
  type BiffRecord,
  type BiffSstEntry,
} from "./biff"

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

const XLS_MAX_COL_INDEX = 255

/** Whether a CFB container holds a BIFF Workbook stream (.xls). */
export function looksLikeXls(streams: Map<string, Uint8Array>): boolean {
  return streams.has("Workbook") || streams.has("Book")
}

/** Read a BIFF8 .xls workbook into the standard {@link Workbook} model. */
export async function readXls(
  input: Uint8Array | ArrayBuffer | ReadableStream<Uint8Array>,
  options?: ReadOptions,
): Promise<Workbook> {
  const data = await readInputToUint8Array(input, options?.maxInputBytes)
  let streams: Map<string, Uint8Array>
  try {
    streams = readCfb(data)
  } catch (err) {
    throw new ParseError("Failed to open XLS: not a valid OLE2 container", undefined, {
      cause: err,
    })
  }
  const stream = streams.get("Workbook") ?? streams.get("Book")
  if (!stream) throw new ParseError("Invalid XLS: missing Workbook stream")

  // Record parsing reads many length-prefixed binary fields; a truncated or
  // hostile file can make DataView accessors throw a raw RangeError. Wrap the
  // whole pass so malformed input surfaces as the library's ParseError.
  try {
    return parseWorkbookRecords(stream, options)
  } catch (err) {
    if (err instanceof ParseError) throw err
    throw new ParseError("Failed to parse XLS workbook (malformed or truncated)", undefined, {
      cause: err,
    })
  }
}

function parseWorkbookRecords(stream: Uint8Array, options?: ReadOptions): Workbook {
  const records = parseRecords(stream)

  // ── BIFF version gate ──
  // The first record is the workbook globals BOF; its first u16 is the BIFF
  // version (0x0600 = BIFF8, 0x0500 = BIFF5/7). The Flyfish fork keeps its
  // legacy BIFF5/7 acceptance path for older WPS and Excel exports. Unknown
  // versions still fail explicitly instead of being parsed as BIFF8.
  const bof = records[0]
  if (!bof || bof.id !== SID.BOF) {
    throw new ParseError("Invalid XLS: missing BOF record at start of Workbook stream")
  }
  let biffVersion = 0x0600
  if (bof.data.length >= 2) {
    biffVersion = new Reader(bof.data).u16()
    if (biffVersion !== 0x0600 && biffVersion !== 0x0500) {
      throw new ParseError(
        `Unsupported XLS version (BIFF 0x${biffVersion.toString(16)}). ` +
          "Only BIFF5/7 and BIFF8 workbooks are supported; re-save the file as .xlsx.",
      )
    }
  }

  const offsetToIndex = new Map<number, number>()
  for (let i = 0; i < records.length; i++) offsetToIndex.set(records[i].offset, i)

  // ── Globals substream (records[0] = BOF … first EOF) ──
  let date1904 = options?.dateSystem === "1904"
  const xfFmtIds: number[] = []
  const fmtCodes = new Map<number, string>()
  const boundSheets: Array<{ name: string; pos: number; hidden: boolean; veryHidden: boolean }> = []
  const sst: BiffSstEntry[] = []
  const fontRecords: Uint8Array[] = []
  const xfRecords: Uint8Array[] = []
  let palette = [...DEFAULT_INDEXED_PALETTE]

  let gi = 0
  for (; gi < records.length; gi++) {
    const rec = records[gi]
    if (rec.id === SID.EOF) {
      gi++
      break
    }
    switch (rec.id) {
      case SID.DATEMODE: {
        if (!options?.dateSystem || options.dateSystem === "auto") {
          date1904 = new Reader(rec.data).u16() === 1
        }
        break
      }
      case SID.FORMAT: {
        const r = new Reader(rec.data)
        const ifmt = r.u16()
        fmtCodes.set(ifmt, readXLString(r))
        break
      }
      case SID.XF: {
        const r = new Reader(rec.data)
        r.u16() // ifnt
        xfFmtIds.push(r.u16()) // ifmt
        if (biffVersion === 0x0600) xfRecords.push(rec.data)
        break
      }
      case SID.FONT: {
        if (biffVersion === 0x0600) fontRecords.push(rec.data)
        break
      }
      case SID.PALETTE: {
        if (biffVersion === 0x0600) palette = parsePalette(rec.data)
        break
      }
      case SID.BOUNDSHEET: {
        const r = new Reader(rec.data)
        const pos = r.u32()
        const state = r.u8()
        r.u8() // dt (sheet type)
        boundSheets.push({
          name: readShortString(r),
          pos,
          hidden: state === 1,
          veryHidden: state === 2,
        })
        break
      }
      case SID.SST: {
        const blocks: Uint8Array[] = [rec.data]
        // Gather trailing CONTINUE records belonging to the SST.
        for (let j = gi + 1; j < records.length; j++) {
          if (records[j].id !== SID.CONTINUE) break
          blocks.push(records[j].data)
        }
        // A loop, not `push(...parseSstEntries(blocks))`: spreading an array as
        // arguments puts one stack slot per element, so a workbook with a
        // few hundred thousand shared strings — an ordinary large .xls —
        // threw `RangeError: Maximum call stack size exceeded`, which the
        // caller reported as "malformed or truncated".
        for (const entry of parseSstEntries(blocks)) sst.push(entry)
        break
      }
      default:
        break
    }
  }

  const dateXf = xfFmtIds.map((id) => {
    if (isBuiltinDateFormatId(id)) return true
    const code = fmtCodes.get(id)
    return code ? isDateFormat(code) : false
  })

  const isDate = (ixfe: number): boolean => dateXf[ixfe] === true
  const fonts =
    biffVersion === 0x0600 && options?.readStyles ? parseBiffFonts(fontRecords, palette) : undefined
  const styles = fonts ? parseBiffStyles(fonts, xfRecords, fmtCodes, palette) : undefined

  // ── Sheet substreams ──
  const sheets: Sheet[] = []
  for (const bs of boundSheets) {
    const startIdx = offsetToIndex.get(bs.pos)
    if (startIdx === undefined) {
      sheets.push({ name: bs.name, rows: [] })
      continue
    }
    sheets.push(
      parseSheet(
        records,
        startIdx,
        bs.name,
        sst,
        isDate,
        date1904,
        options?.maxTotalCells ?? MAX_TOTAL_CELLS,
        styles,
        fonts,
        bs.hidden,
        bs.veryHidden,
      ),
    )
  }

  // Some minimal BIFF5/7 producers omit BOUNDSHEET for an otherwise valid,
  // empty single-sheet workbook. Preserve the fork's established result shape.
  if (sheets.length === 0 && biffVersion === 0x0500) sheets.push({ name: "Sheet1", rows: [] })

  return { sheets }
}

function parseSheet(
  records: BiffRecord[],
  startIdx: number,
  name: string,
  sst: BiffSstEntry[],
  isDate: (ixfe: number) => boolean,
  date1904: boolean,
  cellLimit: number,
  styles?: CellStyle[],
  fonts?: Array<FontStyle | undefined>,
  hidden = false,
  veryHidden = false,
): Sheet {
  const rows: CellValue[][] = []
  const merges: MergeRange[] = []
  const cells = styles ? new Map<string, Cell>() : undefined
  const columns: NonNullable<Sheet["columns"]> = []
  const rowDefs = new Map<number, RowDef>()
  const sheetFormat: NonNullable<Sheet["sheetFormat"]> = {}
  let dimensionLastRow: number | undefined
  let dimensionLastCol: number | undefined

  // BIFF row/col are u16, so each is bounded at 65,535 on its own — but
  // their product is not, and `rows` is a dense rectangle. 65,535 rows of
  // 65,536 slots is 4.3e9 allocations from a few hundred KB of input.
  // The XLSB reader already guards its coordinates this way; this one did
  // not. See #363.
  let widestCol = 0
  const setCell = (
    row: number,
    col: number,
    value: CellValue,
    ixfe?: number,
    richText?: RichTextRun[],
  ): void => {
    if (row < 0 || row > MAX_ROW_INDEX) {
      throw new ParseError(
        `Cell row ${row} is outside the supported sheet bounds (max ${MAX_ROW_INDEX + 1})`,
      )
    }
    if (col < 0 || col > MAX_COL_INDEX) {
      throw new ParseError(
        `Cell column ${col} is outside the supported sheet bounds (max ${MAX_COL_INDEX + 1})`,
      )
    }
    if (col >= widestCol) widestCol = col + 1
    const boundingBox = Math.max(rows.length, row + 1) * widestCol
    if (boundingBox > cellLimit) {
      throw new ParseError(
        `Sheet spans ${boundingBox} cells, over the ${cellLimit} limit. ` +
          "Raise `maxTotalCells` if the sheet really is this large.",
      )
    }
    let r = rows[row]
    if (!r) r = rows[row] = []
    while (r.length < col) r.push(null)
    r[col] = value
    if (cells) {
      const cell: Cell = { value, type: cellTypeOf(value) }
      const style = ixfe === undefined ? undefined : styles?.[ixfe]
      if (style && Object.keys(style).length > 0) cell.style = style
      if (richText && richText.length > 0) {
        cell.type = "richText"
        cell.richText = richText
      }
      cells.set(`${row},${col}`, cell)
    }
  }
  const numeric = (row: number, col: number, ixfe: number, n: number): void => {
    setCell(row, col, isDate(ixfe) ? serialToDate(n, date1904) : n, ixfe)
  }

  for (let i = startIdx + 1; i < records.length; i++) {
    const rec = records[i]
    if (rec.id === SID.EOF) break
    const r = new Reader(rec.data)
    switch (rec.id) {
      case SID.LABELSST: {
        const row = r.u16(),
          col = r.u16()
        const ixfe = r.u16()
        const entry = sst[r.u32()]
        const value = entry?.text ?? ""
        setCell(row, col, value, ixfe, entry ? biffRichText(entry, fonts) : undefined)
        break
      }
      case SID.RK: {
        const row = r.u16(),
          col = r.u16(),
          ixfe = r.u16()
        numeric(row, col, ixfe, decodeRk(r.u32()))
        break
      }
      case SID.NUMBER: {
        const row = r.u16(),
          col = r.u16(),
          ixfe = r.u16()
        numeric(row, col, ixfe, r.f64())
        break
      }
      case SID.MULRK: {
        const row = r.u16()
        const colFirst = r.u16()
        const count = (rec.data.length - 6) / 6
        for (let k = 0; k < count; k++) {
          const ixfe = r.u16()
          numeric(row, colFirst + k, ixfe, decodeRk(r.u32()))
        }
        break
      }
      case SID.BOOLERR: {
        const row = r.u16(),
          col = r.u16()
        const ixfe = r.u16()
        const val = r.u8()
        const isError = r.u8() === 1
        setCell(row, col, isError ? (ERROR_TEXT[val] ?? "#ERR!") : val !== 0, ixfe)
        break
      }
      case SID.LABEL: {
        const row = r.u16(),
          col = r.u16()
        const ixfe = r.u16()
        setCell(row, col, readXLString(r), ixfe)
        break
      }
      case SID.BLANK: {
        const row = r.u16(),
          col = r.u16(),
          ixfe = r.u16()
        if (insideDeclaredDimensions(row, col, dimensionLastRow, dimensionLastCol)) {
          setCell(row, col, null, ixfe)
        }
        break
      }
      case SID.MULBLANK: {
        const row = r.u16()
        const colFirst = r.u16()
        const count = Math.max(0, (rec.data.length - 6) / 2)
        for (let k = 0; k < count; k++) {
          const col = colFirst + k
          const ixfe = r.u16()
          if (insideDeclaredDimensions(row, col, dimensionLastRow, dimensionLastCol)) {
            setCell(row, col, null, ixfe)
          }
        }
        break
      }
      case SID.FORMULA: {
        const row = r.u16(),
          col = r.u16(),
          ixfe = r.u16()
        const b = rec.data.subarray(r.pos, r.pos + 8)
        if (b[6] === 0xff && b[7] === 0xff) {
          const kind = b[0]
          if (kind === 1)
            setCell(row, col, b[2] !== 0, ixfe) // boolean
          else if (kind === 2)
            setCell(row, col, ERROR_TEXT[b[2]] ?? "#ERR!", ixfe) // error
          else if (kind === 0) {
            // string: value is in the following STRING record
            const next = records[i + 1]
            if (next && next.id === SID.STRING)
              setCell(row, col, readXLString(new Reader(next.data)), ixfe)
          } else if (kind === 3) {
            // A cached blank is not a worksheet value. Keep a style-only
            // cell when styles were requested, but preserve the historical
            // sparse rows shape for value-only reads.
            if (cells) setCell(row, col, null, ixfe)
          }
          // kind === 3 → blank/empty
        } else {
          const num = new DataView(b.buffer, b.byteOffset, 8).getFloat64(0, true)
          numeric(row, col, ixfe, num)
        }
        break
      }
      case SID.DIMENSIONS: {
        if (rec.data.length >= 12) {
          r.u32() // first used row
          dimensionLastRow = Math.max(0, r.u32() - 1)
          r.u16() // first used column
          dimensionLastCol = Math.max(0, r.u16() - 1)
        }
        break
      }
      case SID.DEFAULTROWHEIGHT: {
        if (rec.data.length >= 4) {
          const flags = r.u16()
          sheetFormat.zeroHeight = Boolean(flags & 0x0002)
          sheetFormat.defaultRowHeight = r.u16() / 20
        }
        break
      }
      case SID.DEFCOLWIDTH: {
        if (rec.data.length >= 2) sheetFormat.defaultColWidth = r.u16()
        break
      }
      case SID.STANDARDWIDTH: {
        if (rec.data.length >= 2) sheetFormat.defaultColWidth = r.u16() / 256
        break
      }
      case SID.COLINFO: {
        if (rec.data.length >= 10) {
          const first = Math.min(r.u16(), XLS_MAX_COL_INDEX)
          const last = Math.min(r.u16(), XLS_MAX_COL_INDEX)
          const width = r.u16() / 256
          const ixfe = r.u16()
          const flags = r.u16()
          for (let col = first; col <= last; col++) {
            columns[col] = {
              width,
              ...(styles?.[ixfe] && Object.keys(styles[ixfe]).length > 0
                ? { style: styles[ixfe] }
                : {}),
              ...(flags & 0x0001 ? { hidden: true } : {}),
              ...((flags >> 8) & 0x7 ? { outlineLevel: (flags >> 8) & 0x7 } : {}),
              ...(flags & 0x1000 ? { collapsed: true } : {}),
            }
          }
        }
        break
      }
      case SID.ROW: {
        if (rec.data.length >= 16) {
          const row = r.u16()
          r.skip(4) // first and last stored column
          const heightTwips = r.u16()
          r.skip(4)
          const flags = r.u32()
          const ixfe = (flags >>> 16) & 0x0fff
          rowDefs.set(row, {
            height: heightTwips / 20,
            customHeight: Boolean(flags & 0x0040),
            ...(flags & 0x0020 ? { hidden: true } : {}),
            ...(flags & 0x0010 ? { collapsed: true } : {}),
            ...(flags & 0x0007 ? { outlineLevel: flags & 0x0007 } : {}),
            ...(flags & 0x0080 && styles?.[ixfe] && Object.keys(styles[ixfe]).length > 0
              ? { style: styles[ixfe] }
              : {}),
          })
        }
        break
      }
      case SID.MERGECELLS: {
        const cmcs = r.u16()
        for (let k = 0; k < cmcs; k++) {
          const rwFirst = r.u16(),
            rwLast = r.u16(),
            colFirst = r.u16(),
            colLast = r.u16()
          merges.push({ startRow: rwFirst, endRow: rwLast, startCol: colFirst, endCol: colLast })
        }
        break
      }
      default:
        break
    }
  }

  // `rows` is a dense rectangle — the bounding-box guard above is sized
  // on that assumption and `CellValue` has no `undefined` member — but
  // `setCell` only pads a row up to its *own* last written column, and
  // never allocates a row with no cell records at all. So a sheet came
  // back ragged, and a row Excel left empty came back as a hole rather
  // than a row. `readXlsx` normalizes at the end of its parse; this and
  // the XLSB reader did not. See #494.
  densify(rows, widestCol)

  const sheet: Sheet = { name, rows }
  if (merges.length > 0) sheet.merges = merges
  if (cells?.size) sheet.cells = cells
  if (columns.some(Boolean)) sheet.columns = columns
  if (rowDefs.size) sheet.rowDefs = rowDefs
  if (Object.keys(sheetFormat).length > 0) sheet.sheetFormat = sheetFormat
  if (hidden) sheet.hidden = true
  if (veryHidden) sheet.veryHidden = true
  return sheet
}

function insideDeclaredDimensions(
  row: number,
  col: number,
  lastRow: number | undefined,
  lastCol: number | undefined,
): boolean {
  return (lastRow === undefined || row <= lastRow) && (lastCol === undefined || col <= lastCol)
}

function cellTypeOf(value: CellValue): CellType {
  if (value === null) return "empty"
  if (value instanceof Date) return "date"
  if (typeof value === "number") return "number"
  if (typeof value === "boolean") return "boolean"
  if (/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERR!)$/.test(value)) return "error"
  return "string"
}

function parsePalette(data: Uint8Array): string[] {
  const palette = [...DEFAULT_INDEXED_PALETTE]
  if (data.length < 2) return palette
  const r = new Reader(data)
  const count = Math.min(r.u16(), Math.floor(r.remaining() / 4), 56)
  for (let index = 0; index < count; index++) {
    const red = r.u8()
    const green = r.u8()
    const blue = r.u8()
    r.u8()
    palette[index + 8] = [red, green, blue]
      .map((value) => value.toString(16).padStart(2, "0").toUpperCase())
      .join("")
  }
  return palette
}

function parseBiffFonts(
  fontRecords: Uint8Array[],
  palette: string[],
): Array<FontStyle | undefined> {
  const fonts: Array<FontStyle | undefined> = []
  for (const record of fontRecords) {
    // BIFF reserves font index 4 and emits no FONT record for it.
    if (fonts.length === 4) fonts.push(undefined)
    fonts.push(parseBiffFont(record, palette))
  }
  return fonts
}

function parseBiffStyles(
  fonts: Array<FontStyle | undefined>,
  xfRecords: Uint8Array[],
  formats: Map<number, string>,
  palette: string[],
): CellStyle[] {
  return xfRecords.map((record) => parseBiffXf(record, fonts, formats, palette))
}

function biffRichText(
  entry: BiffSstEntry,
  fonts?: Array<FontStyle | undefined>,
): RichTextRun[] | undefined {
  if (!fonts || !entry.runs?.length || entry.text.length === 0) return undefined

  const byStart = new Map<number, number>()
  for (const run of entry.runs) {
    if (run.start < 0 || run.start > entry.text.length) continue
    byStart.set(run.start, run.fontIndex)
  }
  if (byStart.size === 0) return undefined

  const starts = [...byStart.keys()].sort((left, right) => left - right)
  if (starts[0] !== 0) starts.unshift(0)
  const richText: RichTextRun[] = []
  for (let index = 0; index < starts.length; index++) {
    const start = starts[index]
    const end = starts[index + 1] ?? entry.text.length
    if (end <= start) continue
    const text = entry.text.slice(start, end)
    const fontIndex = byStart.get(start)
    const font = fontIndex === undefined ? undefined : fonts[fontIndex]
    richText.push(font ? { text, font } : { text })
  }
  return richText.length > 0 ? richText : undefined
}

function parseBiffFont(data: Uint8Array, palette: string[]): FontStyle {
  if (data.length < 14) return {}
  const r = new Reader(data)
  const heightTwips = r.u16()
  const flags = r.u16()
  const colorIndex = r.u16()
  const weight = r.u16()
  const escapement = r.u16()
  const underline = r.u8()
  const family = r.u8()
  const charset = r.u8()
  r.u8()
  const font: FontStyle = {
    size: heightTwips / 20,
    ...(flags & 0x0002 ? { italic: true } : {}),
    ...(flags & 0x0008 ? { strikethrough: true } : {}),
    ...(weight >= 700 ? { bold: true } : {}),
    ...(underline ? { underline: biffUnderline(underline) } : {}),
    ...(escapement === 1 ? { vertAlign: "superscript" as const } : {}),
    ...(escapement === 2 ? { vertAlign: "subscript" as const } : {}),
    ...(family ? { family } : {}),
    ...(charset ? { charset } : {}),
  }
  const color = colorFromPalette(colorIndex, palette)
  if (color) font.color = color
  if (r.remaining() >= 2) font.name = readShortString(r)
  return font
}

function biffUnderline(value: number): FontStyle["underline"] {
  if (value === 2) return "double"
  if (value === 0x21) return "singleAccounting"
  if (value === 0x22) return "doubleAccounting"
  return "single"
}

function parseBiffXf(
  data: Uint8Array,
  fonts: Array<FontStyle | undefined>,
  formats: Map<number, string>,
  palette: string[],
): CellStyle {
  if (data.length < 6) return {}
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  const fontIndex = view.getUint16(0, true)
  const formatIndex = view.getUint16(2, true)
  const protection = view.getUint16(4, true)
  const style: CellStyle = {}
  // Font zero is the workbook's Normal/default font. Rendering it on every
  // cell is redundant and turns an otherwise sparse workbook into a fully
  // styled grid; non-default font records remain explicit.
  const font = fontIndex === 0 ? undefined : fonts[fontIndex]
  if (font) style.font = font
  const numFmt = formats.get(formatIndex) ?? BUILTIN_NUM_FMTS[formatIndex]
  if (formatIndex !== 0 && numFmt) style.numFmt = numFmt
  const locked = Boolean(protection & 0x0001)
  const formulaHidden = Boolean(protection & 0x0002)
  if (!locked || formulaHidden) style.protection = { locked, hidden: formulaHidden }
  if (data.length < 20) return style

  const alignmentByte = view.getUint8(6)
  const rotation = view.getUint8(7)
  const indentByte = view.getUint8(8)
  const alignment = parseBiffAlignment(alignmentByte, rotation, indentByte)
  if (Object.keys(alignment).length > 0) style.alignment = alignment

  const border1 = view.getUint32(10, true)
  const border2 = view.getUint32(14, true)
  const patternColors = view.getUint16(18, true)
  const border = {
    left: biffBorderSide(border1 & 0x0f, (border1 >>> 16) & 0x7f, palette),
    right: biffBorderSide((border1 >>> 4) & 0x0f, (border1 >>> 23) & 0x7f, palette),
    top: biffBorderSide((border1 >>> 8) & 0x0f, border2 & 0x7f, palette),
    bottom: biffBorderSide((border1 >>> 12) & 0x0f, (border2 >>> 7) & 0x7f, palette),
    diagonal: biffBorderSide((border2 >>> 21) & 0x0f, (border2 >>> 14) & 0x7f, palette),
    diagonalUp: Boolean(border1 & 0x40000000),
    diagonalDown: Boolean(border1 & 0x80000000),
  }
  if (Object.values(border).some(Boolean)) style.border = border

  const pattern = biffFillPattern((border2 >>> 26) & 0x3f)
  if (pattern !== "none") {
    const fgColor = colorFromPalette(patternColors & 0x7f, palette)
    const bgColor = colorFromPalette((patternColors >>> 7) & 0x7f, palette)
    style.fill = {
      type: "pattern",
      pattern,
      ...(fgColor ? { fgColor } : {}),
      ...(bgColor ? { bgColor } : {}),
    }
  }
  return style
}

function parseBiffAlignment(value: number, rotation: number, indentByte: number): AlignmentStyle {
  const horizontal: NonNullable<AlignmentStyle["horizontal"]>[] = [
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "centerContinuous",
    "distributed",
  ]
  const vertical: NonNullable<AlignmentStyle["vertical"]>[] = [
    "top",
    "center",
    "bottom",
    "justify",
    "distributed",
  ]
  const readingOrder = (indentByte >>> 6) & 0x03
  const horizontalValue = horizontal[value & 0x07]
  const verticalValue = vertical[(value >>> 4) & 0x07]
  return {
    ...(horizontalValue && horizontalValue !== "general" ? { horizontal: horizontalValue } : {}),
    ...(verticalValue && verticalValue !== "bottom" ? { vertical: verticalValue } : {}),
    ...(value & 0x08 ? { wrapText: true } : {}),
    ...(indentByte & 0x10 ? { shrinkToFit: true } : {}),
    ...(indentByte & 0x0f ? { indent: indentByte & 0x0f } : {}),
    ...(rotation ? { textRotation: rotation } : {}),
    ...(readingOrder === 1 ? { readingOrder: "ltr" as const } : {}),
    ...(readingOrder === 2 ? { readingOrder: "rtl" as const } : {}),
  }
}

function biffBorderSide(
  styleIndex: number,
  colorIndex: number,
  palette: string[],
): BorderSide | undefined {
  const styles: Array<BorderLineStyle | undefined> = [
    undefined,
    "thin",
    "medium",
    "dashed",
    "dotted",
    "thick",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
  ]
  const borderStyle = styles[styleIndex]
  if (!borderStyle) return undefined
  const color = colorFromPalette(colorIndex, palette)
  return { style: borderStyle, ...(color ? { color } : {}) }
}

function biffFillPattern(value: number): FillPattern {
  const patterns: FillPattern[] = [
    "none",
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625",
  ]
  return patterns[value] ?? "none"
}

function colorFromPalette(index: number, palette: string[]): Color | undefined {
  const rgb = palette[index]
  return rgb ? { rgb } : undefined
}

// ── String helpers ───────────────────────────────────────────────────

/** XLUnicodeString: u16 char count + 1 grbit byte + chars. */
function readXLString(r: Reader): string {
  const cch = r.u16()
  return readChars(r, cch)
}

/** ShortXLUnicodeString: u8 char count + 1 grbit byte + chars. */
function readShortString(r: Reader): string {
  const cch = r.u8()
  return readChars(r, cch)
}

function readChars(r: Reader, cch: number): string {
  const grbit = r.u8()
  const compressed = (grbit & 0x01) === 0
  let s = ""
  for (let i = 0; i < cch; i++) s += String.fromCharCode(compressed ? r.u8() : r.u16())
  return s
}

/**
 * Fill a sparsely-built row array out to a rectangle.
 *
 * Two separate holes, both from building rows only where cells landed: a
 * row index never touched is `undefined` — which `CellValue` cannot
 * express — and a row that ended early is shorter than the sheet. See
 * #494.
 */
export function densify(rows: CellValue[][], width: number): void {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] ?? (rows[r] = [])
    while (row.length < width) row.push(null)
  }
}
