// ── XLSB Reader ──────────────────────────────────────────────────────
// Reads Excel Binary Workbook files (.xlsb). XLSB uses the same OPC/ZIP
// container layout as XLSX, but workbook and worksheet bodies are BIFF12
// binary parts instead of XML parts.

import type { Cell, CellType, CellValue, MergeRange, ReadInput, ReadOptions, Sheet, Workbook } from "../_types"
import { EncryptedFileError, ParseError, ZipError } from "../errors"
import { isDateFormat, serialToDate } from "../_date"
import { decodeBiffFormula } from "../xls/formula"
import { isOle2Container, readInputToUint8Array } from "../_input"
import { ZipReader } from "../zip/reader"
import { parseContentTypes } from "../xlsx/content-types"
import { parseRelationships } from "../xlsx/relationships"
import { parseCoreProperties, parseAppProperties, parseCustomProperties } from "../xlsx/doc-props-reader"
import type { Relationship } from "../xlsx/relationships"
import {
  BRT_BUNDLE_SH,
  BRT_CELL_BLANK,
  BRT_CELL_BOOL,
  BRT_CELL_ERROR,
  BRT_CELL_ISST,
  BRT_CELL_REAL,
  BRT_CELL_RK,
  BRT_CELL_ST,
  BRT_FMLA_BOOL,
  BRT_FMLA_ERROR,
  BRT_FMLA_NUM,
  BRT_FMLA_STRING,
  BRT_BEGIN_CELL_XFS,
  BRT_END_CELL_XFS,
  BRT_FMT,
  BRT_HLINK,
  BRT_MERGE_CELL,
  BRT_ROW_HDR,
  BRT_SST_ITEM,
  BRT_WB_PROP,
  BRT_XF,
  decodeError,
  decodeRk,
  f64,
  readCellHeader,
  readXlsbRecords,
  readXlsbWideString,
  u16,
  u32,
} from "./records"

const NS_TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"

function matchesRelType(rel: string, type: string): boolean {
  return rel === `${NS_TRANSITIONAL}/${type}` || rel === `${NS_STRICT}/${type}` || rel.endsWith("/" + type)
}

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data)
}

function dirname(path: string): string {
  const idx = path.lastIndexOf("/")
  return idx === -1 ? "" : path.slice(0, idx)
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1)
  const baseParts = base.split("/").filter(Boolean)
  const targetParts = target.split("/").filter(Boolean)
  for (const part of targetParts) {
    if (part === "..") baseParts.pop()
    else if (part !== ".") baseParts.push(part)
  }
  return baseParts.join("/")
}

interface XlsbSheetInfo {
  name: string
  index: number
  rId: string
  state: "visible" | "hidden" | "veryHidden"
}

interface XlsbWorkbookInfo {
  sheets: XlsbSheetInfo[]
  dateSystem: "1900" | "1904"
}

interface ParsedXlsbStyles {
  formats: Map<number, string>
  cellXfs: number[]
}

interface ParsedXlsbHLink {
  range: MergeRange
  target: string
  location?: string
  tooltip?: string
}

/** Read an Excel Binary Workbook (.xlsb) and return a Workbook. */
export async function readXlsb(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = await readInputToUint8Array(input)
  if (isOle2Container(data)) {
    throw new EncryptedFileError("xlsb")
  }

  let zip: ZipReader
  try {
    zip = new ZipReader(data)
  } catch (err) {
    if (err instanceof ZipError) throw err
    throw new ParseError("Failed to open XLSB file: not a valid ZIP archive", undefined, { cause: err })
  }

  if (!zip.has("[Content_Types].xml")) {
    throw new ParseError("Invalid XLSB: missing [Content_Types].xml")
  }
  const contentTypesXml = decodeUtf8(await zip.extract("[Content_Types].xml"))
  parseContentTypes(contentTypesXml)

  if (!zip.has("_rels/.rels")) {
    throw new ParseError("Invalid XLSB: missing _rels/.rels")
  }
  const rootRels = parseRelationships(decodeUtf8(await zip.extract("_rels/.rels")))
  const workbookRel = rootRels.find((r) => matchesRelType(r.type, "officeDocument"))
  if (!workbookRel) {
    throw new ParseError("Invalid XLSB: cannot find workbook relationship in _rels/.rels")
  }

  const workbookPath = workbookRel.target.startsWith("/") ? workbookRel.target.slice(1) : workbookRel.target
  if (!zip.has(workbookPath)) throw new ParseError(`Invalid XLSB: missing workbook at ${workbookPath}`)

  const workbookDir = dirname(workbookPath)
  const workbookRelsPath = workbookDir
    ? `${workbookDir}/_rels/${workbookPath.slice(workbookDir.length + 1)}.rels`
    : `_rels/${workbookPath}.rels`

  let workbookRels: Relationship[] = []
  if (zip.has(workbookRelsPath)) {
    workbookRels = parseRelationships(decodeUtf8(await zip.extract(workbookRelsPath)))
  }

  const workbookInfo = parseWorkbookBin(await zip.extract(workbookPath))
  const sheetRelMap = new Map<string, string>()
  for (const rel of workbookRels) {
    if (matchesRelType(rel.type, "worksheet")) sheetRelMap.set(rel.id, resolvePath(workbookDir, rel.target))
  }

  let sharedStrings: string[] = []
  const ssRel = workbookRels.find((r) => matchesRelType(r.type, "sharedStrings"))
  if (ssRel) {
    const ssPath = resolvePath(workbookDir, ssRel.target)
    if (zip.has(ssPath)) sharedStrings = parseSharedStringsBin(await zip.extract(ssPath))
  }

  let styles: ParsedXlsbStyles | undefined
  const stylesRel = workbookRels.find((r) => matchesRelType(r.type, "styles"))
  if (stylesRel) {
    const stylesPath = resolvePath(workbookDir, stylesRel.target)
    if (zip.has(stylesPath)) styles = parseStylesBin(await zip.extract(stylesPath))
  }

  const sheetsToRead = filterSheets(workbookInfo.sheets, options?.sheets)
  const sheets: Sheet[] = []
  for (const info of sheetsToRead) {
    const wsPath = sheetRelMap.get(info.rId)
    if (!wsPath || !zip.has(wsPath)) {
      throw new ParseError(`Invalid XLSB: missing worksheet file for sheet "${info.name}"`)
    }
    const sheet = parseWorksheetBin(await zip.extract(wsPath), info.name, sharedStrings, workbookInfo, styles, options)
    if (info.state === "hidden") sheet.hidden = true
    if (info.state === "veryHidden") sheet.veryHidden = true
    sheets.push(sheet)
  }

  let properties: import("../_types").WorkbookProperties | undefined
  if (zip.has("docProps/core.xml")) {
    const coreProps = parseCoreProperties(decodeUtf8(await zip.extract("docProps/core.xml")))
    if (Object.keys(coreProps).length > 0) properties = { ...coreProps }
  }
  if (zip.has("docProps/app.xml")) {
    const appProps = parseAppProperties(decodeUtf8(await zip.extract("docProps/app.xml")))
    if (Object.keys(appProps).length > 0) properties = { ...properties, ...appProps }
  }
  if (zip.has("docProps/custom.xml")) {
    const customProps = parseCustomProperties(decodeUtf8(await zip.extract("docProps/custom.xml")))
    if (Object.keys(customProps).length > 0) {
      if (!properties) properties = {}
      properties.custom = customProps
    }
  }

  const workbook: Workbook = { sheets, dateSystem: workbookInfo.dateSystem }
  if (properties) workbook.properties = properties
  return workbook
}

function parseWorkbookBin(data: Uint8Array): XlsbWorkbookInfo {
  const sheets: XlsbSheetInfo[] = []
  let dateSystem: "1900" | "1904" = "1900"

  for (const record of readXlsbRecords(data)) {
    if (record.type === BRT_WB_PROP && record.data.length >= 4) {
      // BIFF12 workbook properties carry the date1904 flag in the property
      // bitfield. Excel emits this in BrtWbProp near the top of workbook.bin.
      const props = u32(record.data, 0)
      if ((props & 0x01) !== 0) dateSystem = "1904"
      continue
    }

    if (record.type !== BRT_BUNDLE_SH) continue
    if (record.data.length < 8) continue

    const hidden = u32(record.data, 0)
    const rel = readXlsbWideString(record.data, 8)
    const name = readXlsbWideString(record.data, rel.offset)
    sheets.push({
      name: name.value || `Sheet${sheets.length + 1}`,
      index: sheets.length,
      rId: rel.value,
      state: hidden === 1 ? "hidden" : hidden === 2 ? "veryHidden" : "visible",
    })
  }

  return { sheets, dateSystem }
}

function readOptionalWideString(data: Uint8Array, offset: number): { value: string; offset: number } | null {
  if (offset + 4 > data.length) return null
  const cch = u32(data, offset)
  if (cch > 32767 || offset + 4 + cch * 2 > data.length) return null
  return readXlsbWideString(data, offset)
}

function parseStylesBin(data: Uint8Array): ParsedXlsbStyles {
  const formats = new Map<number, string>()
  const cellXfs: number[] = []
  const allXfs: number[] = []
  let inCellXfs = false
  let sawCellXfsBlock = false

  for (const record of readXlsbRecords(data)) {
    if (record.type === BRT_BEGIN_CELL_XFS) {
      inCellXfs = true
      sawCellXfsBlock = true
      continue
    }
    if (record.type === BRT_END_CELL_XFS) {
      inCellXfs = false
      continue
    }

    if (record.type === BRT_FMT) {
      const id16 = record.data.length >= 2 ? u16(record.data, 0) : -1
      const id32 = record.data.length >= 4 ? u32(record.data, 0) : -1
      const s2 = readOptionalWideString(record.data, 2)
      const s4 = readOptionalWideString(record.data, 4)
      const text = s4?.value || s2?.value || ""
      if (text) formats.set(id32 >= 0 && s4 ? id32 : id16, text)
      continue
    }

    if (record.type === BRT_XF && record.data.length >= 4) {
      // BrtXF follows the BIFF/OOXML style idea of carrying a number-format
      // id near the head of the record. Excel-generated XLSB stores iFmt at
      // offset 2; offset 0 is retained as a fallback for third-party writers.
      const fmtId = u16(record.data, 2)
      allXfs.push(fmtId)
      if (inCellXfs) cellXfs.push(fmtId)
    }
  }

  if (!sawCellXfsBlock && cellXfs.length === 0 && allXfs.length > 0) {
    // Some minimal writers omit the begin/end collection wrappers. StyleXfs
    // normally precede CellXfs, with the first XF being the Normal style; if
    // more than one XF exists, dropping the first aligns style indexes with
    // the cell XF array for common Excel outputs.
    cellXfs.push(...(allXfs.length > 1 ? allXfs.slice(1) : allXfs))
  }

  return { formats, cellXfs }
}

function parseSharedStringsBin(data: Uint8Array): string[] {
  const strings: string[] = []
  for (const record of readXlsbRecords(data)) {
    if (record.type !== BRT_SST_ITEM) continue
    // BrtSSTItem carries an XLWideString payload for plain strings. Rich
    // strings include formatting runs after the text; the leading text still
    // decodes through the same helper and formatting is intentionally ignored
    // in the core CellValue matrix.
    strings.push(readXlsbWideString(record.data, 0).value)
  }
  return strings
}

function parseWorksheetBin(
  data: Uint8Array,
  name: string,
  sharedStrings: string[],
  workbookInfo: XlsbWorkbookInfo,
  styles: ParsedXlsbStyles | undefined,
  options?: ReadOptions,
): Sheet {
  const rows: CellValue[][] = []
  const cells = new Map<string, Cell>()
  const merges: MergeRange[] = []
  const hyperlinks: ParsedXlsbHLink[] = []
  let currentRow = 0
  const maxRows = options?.maxRows ?? 0
  const range = options?.range ? parseRangeRef(options.range) : undefined
  const readStyles = options?.readStyles ?? false

  for (const record of readXlsbRecords(data)) {
    switch (record.type) {
      case BRT_ROW_HDR:
        if (record.data.length >= 4) currentRow = u32(record.data, 0)
        break
      case BRT_CELL_BLANK: {
        const h = readCellHeader(record.data)
        if (!h) break
        setBlank(rows, cells, currentRow, h.col, h.style, styles, readStyles, maxRows, range)
        break
      }
      case BRT_CELL_RK: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length < h.offset + 4) break
        setCell(rows, cells, currentRow, h.col, decodeRk(u32(record.data, h.offset)), "number", h.style, workbookInfo, styles, readStyles, maxRows, range)
        break
      }
      case BRT_CELL_REAL: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length < h.offset + 8) break
        setCell(rows, cells, currentRow, h.col, f64(record.data, h.offset), "number", h.style, workbookInfo, styles, readStyles, maxRows, range)
        break
      }
      case BRT_CELL_BOOL: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length <= h.offset) break
        setCell(rows, cells, currentRow, h.col, record.data[h.offset] === 1, "boolean", h.style, workbookInfo, styles, readStyles, maxRows, range)
        break
      }
      case BRT_CELL_ERROR: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length <= h.offset) break
        setCell(rows, cells, currentRow, h.col, decodeError(record.data[h.offset] ?? 0), "error", h.style, workbookInfo, styles, readStyles, maxRows, range)
        break
      }
      case BRT_CELL_ISST: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length < h.offset + 4) break
        const idx = u32(record.data, h.offset)
        setCell(rows, cells, currentRow, h.col, sharedStrings[idx] ?? "", "string", h.style, workbookInfo, styles, readStyles, maxRows, range)
        break
      }
      case BRT_CELL_ST: {
        const h = readCellHeader(record.data)
        if (!h) break
        const text = readXlsbWideString(record.data, h.offset).value
        setCell(rows, cells, currentRow, h.col, text, "string", h.style, workbookInfo, styles, readStyles, maxRows, range)
        break
      }
      case BRT_FMLA_NUM: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length < h.offset + 8) break
        const value = f64(record.data, h.offset)
        setCell(rows, cells, currentRow, h.col, value, "formula", h.style, workbookInfo, styles, readStyles, maxRows, range, value, decodeFormulaTail(record.data, h.offset + 8, currentRow, h.col, workbookInfo))
        break
      }
      case BRT_FMLA_STRING: {
        const h = readCellHeader(record.data)
        if (!h) break
        const text = readXlsbWideString(record.data, h.offset).value
        setCell(rows, cells, currentRow, h.col, text, "formula", h.style, workbookInfo, styles, readStyles, maxRows, range, text, decodeFormulaTail(record.data, h.offset + 4 + text.length * 2, currentRow, h.col, workbookInfo))
        break
      }
      case BRT_FMLA_BOOL: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length <= h.offset) break
        const value = record.data[h.offset] === 1
        setCell(rows, cells, currentRow, h.col, value, "formula", h.style, workbookInfo, styles, readStyles, maxRows, range, value, decodeFormulaTail(record.data, h.offset + 1, currentRow, h.col, workbookInfo))
        break
      }
      case BRT_FMLA_ERROR: {
        const h = readCellHeader(record.data)
        if (!h || record.data.length <= h.offset) break
        const value = decodeError(record.data[h.offset] ?? 0)
        setCell(rows, cells, currentRow, h.col, value, "formula", h.style, workbookInfo, styles, readStyles, maxRows, range, value, decodeFormulaTail(record.data, h.offset + 1, currentRow, h.col, workbookInfo))
        break
      }
      case BRT_HLINK: {
        const link = parseHLinkBin(record.data)
        if (link) hyperlinks.push(link)
        break
      }
      case BRT_MERGE_CELL:
        if (record.data.length >= 16) {
          merges.push({
            startRow: u32(record.data, 0),
            endRow: u32(record.data, 4),
            startCol: u32(record.data, 8),
            endCol: u32(record.data, 12),
          })
        }
        break
    }
  }

  for (const link of hyperlinks) {
    const row = link.range.startRow
    const col = link.range.startCol
    if (maxRows > 0 && row >= maxRows) continue
    if (!inRange(row, col, range)) continue
    ensureCell(rows, row, col)
    const key = `${row},${col}`
    const existing = cells.get(key) ?? { value: rows[row]?.[col] ?? null, type: "empty" as CellType }
    existing.hyperlink = {
      target: link.target,
      ...(link.location ? { location: link.location } : {}),
      ...(link.tooltip ? { tooltip: link.tooltip } : {}),
    }
    cells.set(key, existing)
  }

  const sheet: Sheet = { name, rows }
  if (cells.size > 0) sheet.cells = cells
  if (merges.length > 0) sheet.merges = merges
  return sheet
}


function numFmtForStyle(styleIndex: number, styles: ParsedXlsbStyles | undefined): string | undefined {
  if (styleIndex < 0) return undefined
  const fmtId = styles?.cellXfs[styleIndex]
  if (fmtId !== undefined) return styles?.formats.get(fmtId) ?? String(fmtId)
  return undefined
}

function decodeFormulaTail(
  data: Uint8Array,
  start: number,
  row: number,
  col: number,
  workbookInfo: XlsbWorkbookInfo,
): string | undefined {
  for (let pos = Math.max(0, start); pos + 2 <= data.length; pos++) {
    const cce16 = u16(data, pos)
    if (cce16 > 0 && cce16 <= data.length - pos - 2) {
      try {
        const formula = decodeBiffFormula(data.subarray(pos + 2, pos + 2 + cce16), {
          currentRow: row,
          currentCol: col,
          sheetNames: workbookInfo.sheets.map((s) => s.name),
        })
        if (formula) return formula
      } catch {
        // Try the next possible alignment.
      }
    }
    if (pos + 4 <= data.length) {
      const cce32 = u32(data, pos)
      if (cce32 > 0 && cce32 < 8192 && cce32 <= data.length - pos - 4) {
        try {
          const formula = decodeBiffFormula(data.subarray(pos + 4, pos + 4 + cce32), {
            currentRow: row,
            currentCol: col,
            sheetNames: workbookInfo.sheets.map((s) => s.name),
          })
          if (formula) return formula
        } catch {
          // Try the next possible alignment.
        }
      }
    }
  }
  return undefined
}

function parseHLinkBin(data: Uint8Array): ParsedXlsbHLink | null {
  if (data.length < 16) return null
  const range: MergeRange = {
    startRow: u32(data, 0),
    endRow: u32(data, 4),
    startCol: u32(data, 8),
    endCol: u32(data, 12),
  }
  const strings: string[] = []
  let pos = 16
  while (pos + 4 <= data.length) {
    const parsed = readOptionalWideString(data, pos)
    if (parsed && parsed.value) {
      strings.push(parsed.value.replace(/\u0000+$/g, ""))
      pos = parsed.offset
    } else {
      pos++
    }
  }
  const target =
    strings.find((s) => /^(https?|ftp|mailto|file):/i.test(s)) ??
    strings.find((s) => /^#/.test(s)) ??
    strings.find((s) => /!|\\|\//.test(s)) ??
    ""
  if (!target) return null
  const link: ParsedXlsbHLink = { range, target: target.startsWith("#") ? "" : target }
  if (target.startsWith("#")) link.location = target.slice(1)
  const tip = strings.find((s) => s && s !== target)
  if (tip) link.tooltip = tip
  return link
}

function setBlank(
  rows: CellValue[][],
  cells: Map<string, Cell>,
  row: number,
  col: number,
  styleIndex: number,
  styles: ParsedXlsbStyles | undefined,
  readStyles: boolean,
  maxRows: number,
  range?: MergeRange,
): void {
  if (maxRows > 0 && row >= maxRows) return
  if (!inRange(row, col, range)) return
  ensureCell(rows, row, col)
  if (readStyles) {
    const numFmt = numFmtForStyle(styleIndex, styles)
    cells.set(`${row},${col}`, { value: null, type: "empty", ...(numFmt ? { style: { numFmt } } : {}) })
  }
}

function setCell(
  rows: CellValue[][],
  cells: Map<string, Cell>,
  row: number,
  col: number,
  rawValue: CellValue,
  type: CellType,
  styleIndex: number,
  workbookInfo: XlsbWorkbookInfo,
  styles: ParsedXlsbStyles | undefined,
  readStyles: boolean,
  maxRows: number,
  range?: MergeRange,
  formulaResult?: CellValue,
  formula?: string,
): void {
  if (maxRows > 0 && row >= maxRows) return
  if (!inRange(row, col, range)) return

  const numFmt = numFmtForStyle(styleIndex, styles)
  let value = rawValue
  let cellType = type
  if (type === "number" && typeof rawValue === "number" && numFmt && isDateFormat(numFmt)) {
    value = serialToDate(rawValue, workbookInfo.dateSystem === "1904")
    cellType = "date"
  }

  ensureCell(rows, row, col)
  rows[row]![col] = value

  const cell: Cell = { value, type: cellType }
  if (type === "formula") {
    cell.type = "formula"
    cell.formula = formula ?? ""
    if (formulaResult !== undefined) cell.formulaResult = formulaResult
  }
  if (readStyles && numFmt) cell.style = { numFmt }
  cells.set(`${row},${col}`, cell)
}

function ensureCell(rows: CellValue[][], row: number, col: number): void {
  while (rows.length <= row) rows.push([])
  const target = rows[row]!
  while (target.length <= col) target.push(null)
}

function parseRangeRef(ref: string): MergeRange {
  const [startRef, endRef] = ref.split(":")
  const start = parseCellRef(startRef ?? "A1")
  const end = parseCellRef(endRef ?? startRef ?? "A1")
  return { startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col }
}

function parseCellRef(ref: string): { row: number; col: number } {
  const clean = ref.replace(/\$/g, "")
  let i = 0
  let col = 0
  while (i < clean.length) {
    const code = clean.charCodeAt(i)
    if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
    else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
    else break
    i++
  }
  const row = Number(clean.slice(i)) - 1
  return { row, col: col - 1 }
}

function inRange(row: number, col: number, range?: MergeRange): boolean {
  if (!range) return true
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}

type SheetFilterInfo = { name: string; index: number; hidden?: boolean; veryHidden?: boolean }

function filterSheets(
  sheets: XlsbSheetInfo[],
  filter: ReadOptions["sheets"] | undefined,
): XlsbSheetInfo[] {
  if (filter === undefined) return sheets
  if (typeof filter === "function") {
    return sheets.filter((s) =>
      filter({ name: s.name, index: s.index, hidden: s.state === "hidden", veryHidden: s.state === "veryHidden" } as SheetFilterInfo),
    )
  }
  const filters = Array.isArray(filter) ? filter : [filter]
  return sheets.filter((s) => filters.some((f) => (typeof f === "number" ? f === s.index : f === s.name)))
}
