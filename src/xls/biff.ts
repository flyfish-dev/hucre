// ── BIFF8 XLS Workbook Reader ────────────────────────────────────────
// Parses the Workbook/Book stream inside legacy XLS CFB containers.
// Implements the BIFF records needed for complete worksheet value reads:
// sheet metadata, shared strings, numbers, RK values, booleans/errors,
// formulas with cached results, labels, merges, row/column metadata, and
// date style inference.

import type { Cell, CellType, CellValue, MergeRange, NamedRange, ReadOptions, Sheet, Workbook } from "../_types"
import { isDateFormat, serialToDate } from "../_date"
import { EncryptedFileError, ParseError } from "../errors"
import { decodeBiffFormula } from "./formula"
import type { FormulaExternSheetRef, FormulaNameRef } from "./formula"

const BIFF_BOF = 0x0809
const BIFF_EOF = 0x000a
const BIFF_CONTINUE = 0x003c
const BIFF_BOUNDSHEET8 = 0x0085
const BIFF_CODEPAGE = 0x0042
const BIFF_DATEMODE = 0x0022
const BIFF_FORMAT = 0x041e
const BIFF_XF = 0x00e0
const BIFF_SST = 0x00fc
const BIFF_FILEPASS = 0x002f
const BIFF_DIMENSIONS = 0x0200
const BIFF_NUMBER = 0x0203
const BIFF_LABEL = 0x0204
const BIFF_BOOLERR = 0x0205
const BIFF_FORMULA = 0x0006
const BIFF_STRING = 0x0207
const BIFF_BLANK = 0x0201
const BIFF_LABELSST = 0x00fd
const BIFF_RK = 0x027e
const BIFF_MULRK = 0x00bd
const BIFF_MULBLANK = 0x00be
const BIFF_MERGECELLS = 0x00e5
const BIFF_ROW = 0x0208
const BIFF_COLINFO = 0x007d
const BIFF_NAME = 0x0018
const BIFF_EXTERNSHEET = 0x0017
const BIFF_HLINK = 0x01b8

interface BiffRecord {
  sid: number
  offset: number
  data: Uint8Array
}

interface BiffSheetInfo {
  name: string
  index: number
  offset: number
  state: "visible" | "hidden" | "veryHidden"
  type: number
}

interface ParsedGlobals {
  sheets: BiffSheetInfo[]
  sharedStrings: string[]
  dateSystem: "1900" | "1904"
  formats: Map<number, string>
  xfs: number[]
  codePage: number
  biffVersion: number
  names: FormulaNameRef[]
  namedRanges: NamedRange[]
  externSheets: FormulaExternSheetRef[]
}

interface PendingFormulaString {
  row: number
  col: number
  xfIndex: number
  formula?: string
}

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

function makeView(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function u16(bytes: Uint8Array, offset: number): number {
  return makeView(bytes).getUint16(offset, true)
}

function u32(bytes: Uint8Array, offset: number): number {
  return makeView(bytes).getUint32(offset, true)
}

function i32(bytes: Uint8Array, offset: number): number {
  return makeView(bytes).getInt32(offset, true)
}

function f64(bytes: Uint8Array, offset: number): number {
  return makeView(bytes).getFloat64(offset, true)
}

function decodeUtf16Le(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-16le").decode(bytes)
  } catch {
    let s = ""
    for (let i = 0; i + 1 < bytes.length; i += 2) s += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8))
    return s
  }
}

function decodeCodePage(bytes: Uint8Array, codePage: number): string {
  const label = codePageToEncoding(codePage)
  try {
    return new TextDecoder(label).decode(bytes)
  } catch {
    // BIFF8 compressed Unicode stores low bytes for Latin text. Latin-1 is a
    // safe fallback for legacy ANSI workbooks when a platform lacks the exact
    // Windows code page decoder.
    let s = ""
    for (const b of bytes) s += String.fromCharCode(b)
    return s
  }
}

function codePageToEncoding(codePage: number): string {
  switch (codePage) {
    case 65001:
      return "utf-8"
    case 1200:
      return "utf-16le"
    case 932:
      return "shift_jis"
    case 936:
      return "gbk"
    case 949:
      return "euc-kr"
    case 950:
      return "big5"
    case 1250:
      return "windows-1250"
    case 1251:
      return "windows-1251"
    case 1252:
    default:
      return "windows-1252"
  }
}

function concatChunks(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, c) => sum + c.length, 0)
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

function readRecords(stream: Uint8Array, startOffset = 0): BiffRecord[] {
  const records: BiffRecord[] = []
  let pos = startOffset
  while (pos + 4 <= stream.length) {
    const sid = u16(stream, pos)
    const len = u16(stream, pos + 2)
    const dataStart = pos + 4
    const dataEnd = dataStart + len
    if (dataEnd > stream.length) {
      throw new ParseError(`Invalid XLS BIFF stream: record 0x${sid.toString(16)} extends past end`)
    }
    records.push({ sid, offset: pos, data: stream.subarray(dataStart, dataEnd) })
    pos = dataEnd
  }
  return records
}

function readBiff8UnicodeString(data: Uint8Array, offset: number, codePage: number): { value: string; offset: number } {
  if (offset + 3 > data.length) return { value: "", offset: data.length }
  const cch = u16(data, offset)
  let pos = offset + 2
  const flags = data[pos++] ?? 0
  const is16 = (flags & 0x01) !== 0
  const hasExtRst = (flags & 0x04) !== 0
  const hasRich = (flags & 0x08) !== 0
  const richRuns = hasRich && pos + 2 <= data.length ? u16(data, pos) : 0
  if (hasRich) pos += 2
  const extSize = hasExtRst && pos + 4 <= data.length ? u32(data, pos) : 0
  if (hasExtRst) pos += 4

  const byteLen = cch * (is16 ? 2 : 1)
  const raw = data.subarray(pos, Math.min(pos + byteLen, data.length))
  pos += byteLen
  const value = is16 ? decodeUtf16Le(raw) : decodeCodePage(raw, codePage)

  // Skip rich formatting runs and phonetic data when present.
  pos += richRuns * 4 + extSize
  return { value, offset: Math.min(pos, data.length) }
}

function readShortBiff8UnicodeString(data: Uint8Array, offset: number, codePage: number): { value: string; offset: number } {
  if (offset + 2 > data.length) return { value: "", offset: data.length }
  const cch = data[offset] ?? 0
  const flags = data[offset + 1] ?? 0
  const is16 = (flags & 0x01) !== 0
  const pos = offset + 2
  const byteLen = cch * (is16 ? 2 : 1)
  const raw = data.subarray(pos, Math.min(pos + byteLen, data.length))
  return {
    value: is16 ? decodeUtf16Le(raw) : decodeCodePage(raw, codePage),
    offset: Math.min(pos + byteLen, data.length),
  }
}

class SstCursor {
  private chunkIndex = 0
  private pos = 0

  constructor(private readonly chunks: Uint8Array[]) {}

  readU8(): number {
    this.ensureChunk()
    return this.chunks[this.chunkIndex]![this.pos++] ?? 0
  }

  readU16(): number {
    const lo = this.readU8()
    const hi = this.readU8()
    return lo | (hi << 8)
  }

  readU32(): number {
    return this.readU8() | (this.readU8() << 8) | (this.readU8() << 16) | (this.readU8() << 24)
  }

  readString(codePage: number): string {
    const cch = this.readU16()
    let flags = this.readU8()
    const hasExtRst = (flags & 0x04) !== 0
    const hasRich = (flags & 0x08) !== 0
    const richRuns = hasRich ? this.readU16() : 0
    const extSize = hasExtRst ? this.readU32() : 0

    let is16 = (flags & 0x01) !== 0
    let out = ""
    let remainingChars = cch

    while (remainingChars > 0) {
      if (this.atChunkEnd()) {
        // When an XLUnicodeString's character payload is split by a CONTINUE
        // record, the continuation starts with a fresh option flag byte that
        // can switch compressed/uncompressed representation.
        this.nextChunk()
        flags = this.readU8()
        is16 = (flags & 0x01) !== 0
      }

      const chunk = this.chunks[this.chunkIndex]!
      const bytesPerChar = is16 ? 2 : 1
      const availableChars = Math.floor((chunk.length - this.pos) / bytesPerChar)
      const takeChars = Math.min(remainingChars, availableChars)
      const byteLen = takeChars * bytesPerChar
      const raw = chunk.subarray(this.pos, this.pos + byteLen)
      this.pos += byteLen
      remainingChars -= takeChars
      out += is16 ? decodeUtf16Le(raw) : decodeCodePage(raw, codePage)

      if (takeChars === 0) this.nextChunk()
    }

    // Rich-text runs and ExtRst blocks may also be continued, but unlike the
    // character payload they do not carry continuation option flags.
    this.skip(richRuns * 4 + extSize)
    return out
  }

  skip(count: number): void {
    let remaining = count
    while (remaining > 0) {
      this.ensureChunk()
      const chunk = this.chunks[this.chunkIndex]!
      const take = Math.min(remaining, chunk.length - this.pos)
      this.pos += take
      remaining -= take
      if (remaining > 0) this.nextChunk()
    }
  }

  private atChunkEnd(): boolean {
    return this.chunkIndex < this.chunks.length && this.pos >= this.chunks[this.chunkIndex]!.length
  }

  private ensureChunk(): void {
    while (this.chunkIndex < this.chunks.length && this.pos >= this.chunks[this.chunkIndex]!.length) {
      this.nextChunk()
    }
    if (this.chunkIndex >= this.chunks.length) {
      throw new ParseError("Invalid XLS SST: unexpected end of continuation records")
    }
  }

  private nextChunk(): void {
    this.chunkIndex++
    this.pos = 0
    if (this.chunkIndex >= this.chunks.length) {
      throw new ParseError("Invalid XLS SST: continuation chain ended early")
    }
  }
}

function parseSst(chunks: Uint8Array[], codePage: number): string[] {
  const cursor = new SstCursor(chunks)
  cursor.readU32() // total string count
  const unique = cursor.readU32()
  const strings: string[] = []
  for (let i = 0; i < unique; i++) {
    strings.push(cursor.readString(codePage))
  }
  return strings
}

function parseBoundSheet(record: Uint8Array, index: number, codePage: number, biff8: boolean): BiffSheetInfo | null {
  if (record.length < 8) return null
  const offset = u32(record, 0)
  const hsState = record[4] ?? 0
  const type = record[5] ?? 0
  const { value: name } = biff8
    ? readShortBiff8UnicodeString(record, 6, codePage)
    : readBiff5ByteString(record, 6, codePage)
  return {
    name,
    index,
    offset,
    state: hsState === 1 ? "hidden" : hsState === 2 ? "veryHidden" : "visible",
    type,
  }
}

function parseFormat(record: Uint8Array, codePage: number): { id: number; value: string } | null {
  if (record.length < 5) return null
  const id = u16(record, 0)
  const { value } = readBiff8UnicodeString(record, 2, codePage)
  return { id, value }
}


function readBiff5ByteString(data: Uint8Array, offset: number, codePage: number): { value: string; offset: number } {
  if (offset >= data.length) return { value: "", offset: data.length }
  const cch = data[offset] ?? 0
  const start = offset + 1
  const end = Math.min(start + cch, data.length)
  return { value: decodeCodePage(data.subarray(start, end), codePage), offset: end }
}

function readBiff8NameText(data: Uint8Array, offset: number, cch: number, codePage: number): { value: string; offset: number } {
  if (offset >= data.length) return { value: "", offset: data.length }
  const flags = data[offset] ?? 0
  const is16 = (flags & 0x01) !== 0
  const start = offset + 1
  const byteLen = cch * (is16 ? 2 : 1)
  const raw = data.subarray(start, Math.min(start + byteLen, data.length))
  return { value: is16 ? decodeUtf16Le(raw) : decodeCodePage(raw, codePage), offset: Math.min(start + byteLen, data.length) }
}

const BUILTIN_NAMES: Record<number, string> = {
  0x00: "Consolidate_Area",
  0x01: "Auto_Open",
  0x02: "Auto_Close",
  0x03: "Extract",
  0x04: "Database",
  0x05: "Criteria",
  0x06: "Print_Area",
  0x07: "Print_Titles",
  0x08: "Recorder",
  0x09: "Data_Form",
  0x0a: "Auto_Activate",
  0x0b: "Auto_Deactivate",
  0x0c: "Sheet_Title",
  0x0d: "_FilterDatabase",
}

function parseExternSheet(record: Uint8Array): FormulaExternSheetRef[] {
  if (record.length < 2) return []
  const count = u16(record, 0)
  const refs: FormulaExternSheetRef[] = []
  let pos = 2
  for (let i = 0; i < count && pos + 6 <= record.length; i++) {
    // The first field is the SupBook index. Local references use the sheet
    // indexes that follow; external-book resolution is intentionally surfaced
    // as a plain reference placeholder by the formula decoder.
    const firstSheet = u16(record, pos + 2)
    const lastSheet = u16(record, pos + 4)
    refs.push({ firstSheet, lastSheet })
    pos += 6
  }
  return refs
}

function parseNameRecord(
  record: Uint8Array,
  globals: Pick<ParsedGlobals, "codePage" | "sheets" | "externSheets" | "names">,
): { name: FormulaNameRef; namedRange?: NamedRange } | null {
  if (record.length < 14) return null
  const flags = u16(record, 0)
  const cch = record[3] ?? 0
  const cce = u16(record, 4)
  const sheetIndex = u16(record, 8)
  const menuLen = record[10] ?? 0
  const descLen = record[11] ?? 0
  const helpLen = record[12] ?? 0
  const statusLen = record[13] ?? 0

  let pos = 14
  let nameText = ""
  if ((flags & 0x0020) !== 0 && cch === 1 && pos < record.length) {
    nameText = BUILTIN_NAMES[record[pos] ?? 0] ?? `_xlnm.Builtin_${record[pos] ?? 0}`
    pos++
  } else {
    const parsed = readBiff8NameText(record, pos, cch, globals.codePage)
    nameText = parsed.value
    pos = parsed.offset
  }

  const formulaBytes = record.subarray(pos, Math.min(pos + cce, record.length))
  const formula = decodeBiffFormula(formulaBytes, {
    currentRow: 0,
    currentCol: 0,
    sheetNames: globals.sheets.map((s) => s.name),
    externSheets: globals.externSheets,
    names: globals.names,
  })
  pos += cce
  pos += menuLen + descLen + helpLen + statusLen

  const name: FormulaNameRef = { name: nameText }
  if (!nameText || !formula) return { name }

  const hidden = (flags & 0x0001) !== 0
  if (hidden && nameText !== "Print_Area" && nameText !== "Print_Titles" && nameText !== "_FilterDatabase") {
    return { name }
  }

  const namedRange: NamedRange = {
    name: nameText,
    range: formula,
  }
  if (sheetIndex > 0) {
    const sheet = globals.sheets[sheetIndex - 1]
    if (sheet) namedRange.scope = sheet.name
  }
  return { name, namedRange }
}

interface ParsedHLink {
  range: MergeRange
  target: string
  location?: string
  tooltip?: string
}

function parseHLink(record: Uint8Array): ParsedHLink | null {
  if (record.length < 28) return null
  const range: MergeRange = {
    startRow: u16(record, 0),
    endRow: u16(record, 2),
    startCol: u16(record, 4),
    endCol: u16(record, 6),
  }

  const strings: string[] = []
  let pos = 28
  while (pos + 4 <= record.length) {
    const cch = u32(record, pos)
    if (cch > 0 && cch < 4096 && pos + 4 + cch * 2 <= record.length) {
      const raw = record.subarray(pos + 4, pos + 4 + cch * 2)
      strings.push(decodeUtf16Le(raw).replace(/\u0000+$/g, ""))
      pos += 4 + cch * 2
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

  const link: ParsedHLink = { range, target: target.startsWith("#") ? "" : target }
  if (target.startsWith("#")) link.location = target.slice(1)
  const tip = strings.find((s) => s && s !== target)
  if (tip) link.tooltip = tip
  return link
}

function parseGlobals(records: BiffRecord[]): ParsedGlobals {
  let codePage = 1252
  let dateSystem: "1900" | "1904" = "1900"
  const sheets: BiffSheetInfo[] = []
  const formats = new Map<number, string>()
  const xfs: number[] = []
  let sharedStrings: string[] = []
  let biffVersion = 0x0600
  const names: FormulaNameRef[] = []
  const namedRanges: NamedRange[] = []
  let externSheets: FormulaExternSheetRef[] = []

  for (let i = 0; i < records.length; i++) {
    const record = records[i]!
    switch (record.sid) {
      case BIFF_BOF:
        if (record.data.length >= 2) biffVersion = u16(record.data, 0)
        break
      case BIFF_FILEPASS:
        throw new EncryptedFileError("xls")
      case BIFF_CODEPAGE:
        if (record.data.length >= 2) codePage = u16(record.data, 0)
        break
      case BIFF_DATEMODE:
        if (record.data.length >= 2) dateSystem = u16(record.data, 0) === 1 ? "1904" : "1900"
        break
      case BIFF_FORMAT: {
        const fmt = parseFormat(record.data, codePage)
        if (fmt) formats.set(fmt.id, fmt.value)
        break
      }
      case BIFF_XF:
        if (record.data.length >= 4) xfs.push(u16(record.data, 2))
        break
      case BIFF_BOUNDSHEET8: {
        const sheet = parseBoundSheet(record.data, sheets.length, codePage, biffVersion >= 0x0600)
        if (sheet) sheets.push(sheet)
        break
      }
      case BIFF_EXTERNSHEET:
        externSheets = parseExternSheet(record.data)
        break
      case BIFF_NAME: {
        const parsed = parseNameRecord(record.data, { codePage, sheets, externSheets, names })
        if (parsed) {
          names.push(parsed.name)
          if (parsed.namedRange) namedRanges.push(parsed.namedRange)
        }
        break
      }
      case BIFF_SST: {
        const chunks = [record.data]
        while (i + 1 < records.length && records[i + 1]!.sid === BIFF_CONTINUE) {
          chunks.push(records[++i]!.data)
        }
        sharedStrings = parseSst(chunks, codePage)
        break
      }
      case BIFF_EOF:
        return { sheets, sharedStrings, dateSystem, formats, xfs, codePage, biffVersion, names, namedRanges, externSheets }
    }
  }

  return { sheets, sharedStrings, dateSystem, formats, xfs, codePage, biffVersion, names, namedRanges, externSheets }
}

function decodeRk(rk: number): number {
  let value: number
  if ((rk & 0x02) !== 0) {
    value = rk >> 2
  } else {
    const buf = new ArrayBuffer(8)
    const view = new DataView(buf)
    view.setUint32(0, 0, true)
    view.setUint32(4, rk & 0xfffffffc, true)
    value = view.getFloat64(0, true)
  }
  if ((rk & 0x01) !== 0) value /= 100
  return value
}

function parseRangeRef(ref: string): MergeRange {
  const [startRef, endRef = startRef] = ref.split(":")
  const start = parseA1(startRef!)
  const end = parseA1(endRef!)
  return { startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col }
}

function parseA1(ref: string): { row: number; col: number } {
  const match = ref.replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/)
  if (!match) return { row: 0, col: 0 }
  let col = 0
  for (const ch of match[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(match[2]) - 1, col: col - 1 }
}

function inRange(row: number, col: number, range?: MergeRange): boolean {
  if (!range) return true
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
}

function numFmtForXf(xfIndex: number, globals: ParsedGlobals): string | undefined {
  if (xfIndex < 0) return undefined
  const fmtId = globals.xfs[xfIndex]
  if (fmtId === undefined) return undefined
  return globals.formats.get(fmtId) ?? String(fmtId)
}

function convertNumber(value: number, xfIndex: number, globals: ParsedGlobals): { value: CellValue; type: CellType } {
  const numFmt = numFmtForXf(xfIndex, globals)
  if (numFmt && isDateFormat(numFmt)) {
    return { value: serialToDate(value, globals.dateSystem === "1904"), type: "date" }
  }
  return { value, type: "number" }
}

function setCell(
  sheet: Sheet,
  row: number,
  col: number,
  value: CellValue,
  type: CellType,
  xfIndex: number,
  globals: ParsedGlobals,
  readStyles: boolean,
): void {
  while (sheet.rows.length <= row) sheet.rows.push([])
  const arr = sheet.rows[row]!
  while (arr.length <= col) arr.push(null)
  arr[col] = value

  const cell: Cell = { value, type }
  if (readStyles) {
    const numFmt = numFmtForXf(xfIndex, globals)
    if (numFmt) cell.style = { numFmt }
  }

  if (!sheet.cells) sheet.cells = new Map()
  sheet.cells.set(`${row},${col}`, cell)
}

function setFormulaCell(
  sheet: Sheet,
  row: number,
  col: number,
  value: CellValue,
  xfIndex: number,
  globals: ParsedGlobals,
  readStyles: boolean,
  formula?: string,
): void {
  while (sheet.rows.length <= row) sheet.rows.push([])
  const arr = sheet.rows[row]!
  while (arr.length <= col) arr.push(null)
  arr[col] = value

  const cell: Cell = { value, type: "formula", formula: formula ?? "", formulaResult: value }
  if (readStyles) {
    const numFmt = numFmtForXf(xfIndex, globals)
    if (numFmt) cell.style = { numFmt }
  }

  if (!sheet.cells) sheet.cells = new Map()
  sheet.cells.set(`${row},${col}`, cell)
}

function setBlank(sheet: Sheet, row: number, col: number, xfIndex: number, globals: ParsedGlobals, readStyles: boolean): void {
  if (!readStyles) return
  setCell(sheet, row, col, null, "empty", xfIndex, globals, true)
}

function parseMergeCells(record: Uint8Array): MergeRange[] {
  if (record.length < 2) return []
  const count = u16(record, 0)
  const merges: MergeRange[] = []
  let pos = 2
  for (let i = 0; i < count && pos + 8 <= record.length; i++) {
    const startRow = u16(record, pos)
    const endRow = u16(record, pos + 2)
    const startCol = u16(record, pos + 4)
    const endCol = u16(record, pos + 6)
    merges.push({ startRow, startCol, endRow, endCol })
    pos += 8
  }
  return merges
}

function shouldReadSheet(sheet: BiffSheetInfo, filter: ReadOptions["sheets"] | undefined): boolean {
  if (filter === undefined) return true
  const info = {
    name: sheet.name,
    index: sheet.index,
    hidden: sheet.state === "hidden",
    veryHidden: sheet.state === "veryHidden",
  }
  if (typeof filter === "function") return filter(info)
  const items = Array.isArray(filter) ? filter : [filter]
  return items.some((item) => (typeof item === "number" ? item === sheet.index : item === sheet.name))
}

function parseWorksheet(stream: Uint8Array, sheetInfo: BiffSheetInfo, globals: ParsedGlobals, options?: ReadOptions): Sheet {
  const sheet: Sheet = { name: sheetInfo.name, rows: [] }
  if (sheetInfo.state === "hidden") sheet.hidden = true
  if (sheetInfo.state === "veryHidden") sheet.veryHidden = true

  const records = readRecords(stream, sheetInfo.offset)
  const readStyles = options?.readStyles ?? false
  const range = options?.range ? parseRangeRef(options.range) : undefined
  const maxRows = options?.maxRows ?? 0
  let pendingFormulaString: PendingFormulaString | undefined
  const hyperlinks: ParsedHLink[] = []

  for (const record of records) {
    if (record.sid === BIFF_EOF) break

    switch (record.sid) {
      case BIFF_FILEPASS:
        throw new EncryptedFileError("xls")
      case BIFF_DIMENSIONS:
        // DIMENSIONS is metadata only; rows are allocated lazily as values arrive.
        break
      case BIFF_NUMBER: {
        if (record.data.length < 14) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        const xf = u16(record.data, 4)
        const converted = convertNumber(f64(record.data, 6), xf, globals)
        setCell(sheet, row, col, converted.value, converted.type, xf, globals, readStyles)
        break
      }
      case BIFF_LABELSST: {
        if (record.data.length < 10) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        const xf = u16(record.data, 4)
        const idx = u32(record.data, 6)
        setCell(sheet, row, col, globals.sharedStrings[idx] ?? "", "string", xf, globals, readStyles)
        break
      }
      case BIFF_LABEL: {
        if (record.data.length < 8) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        const xf = u16(record.data, 4)
        const { value } = globals.biffVersion >= 0x0600
          ? readBiff8UnicodeString(record.data, 6, globals.codePage)
          : readBiff5ByteString(record.data, 6, globals.codePage)
        setCell(sheet, row, col, value, "string", xf, globals, readStyles)
        break
      }
      case BIFF_RK: {
        if (record.data.length < 10) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        const xf = u16(record.data, 4)
        const converted = convertNumber(decodeRk(i32(record.data, 6)), xf, globals)
        setCell(sheet, row, col, converted.value, converted.type, xf, globals, readStyles)
        break
      }
      case BIFF_MULRK: {
        if (record.data.length < 6) break
        const row = u16(record.data, 0)
        if (maxRows > 0 && row >= maxRows) break
        const firstCol = u16(record.data, 2)
        const lastCol = u16(record.data, record.data.length - 2)
        let pos = 4
        for (let col = firstCol; col <= lastCol && pos + 6 <= record.data.length - 2; col++) {
          const xf = u16(record.data, pos)
          const rk = i32(record.data, pos + 2)
          if (inRange(row, col, range)) {
            const converted = convertNumber(decodeRk(rk), xf, globals)
            setCell(sheet, row, col, converted.value, converted.type, xf, globals, readStyles)
          }
          pos += 6
        }
        break
      }
      case BIFF_BOOLERR: {
        if (record.data.length < 8) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        const xf = u16(record.data, 4)
        const raw = record.data[6] ?? 0
        const isError = (record.data[7] ?? 0) !== 0
        setCell(
          sheet,
          row,
          col,
          isError ? (ERROR_TEXT[raw] ?? `#ERR${raw}`) : raw !== 0,
          isError ? "error" : "boolean",
          xf,
          globals,
          readStyles,
        )
        break
      }
      case BIFF_FORMULA: {
        if (record.data.length < 20) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        const xf = u16(record.data, 4)
        const formula = record.data.length >= 22
          ? decodeBiffFormula(record.data.subarray(22, Math.min(22 + u16(record.data, 20), record.data.length)), {
              currentRow: row,
              currentCol: col,
              sheetNames: globals.sheets.map((s) => s.name),
              externSheets: globals.externSheets,
              names: globals.names,
            })
          : ""
        const marker6 = record.data[12]
        const marker7 = record.data[13]
        if (marker6 === 0xff && marker7 === 0xff) {
          const kind = record.data[6] ?? 0
          if (kind === 0) {
            pendingFormulaString = { row, col, xfIndex: xf, formula }
          } else if (kind === 1) {
            setFormulaCell(sheet, row, col, (record.data[8] ?? 0) !== 0, xf, globals, readStyles, formula)
          } else if (kind === 2) {
            const raw = record.data[8] ?? 0
            setFormulaCell(sheet, row, col, ERROR_TEXT[raw] ?? `#ERR${raw}`, xf, globals, readStyles, formula)
          } else {
            setFormulaCell(sheet, row, col, null, xf, globals, readStyles, formula)
          }
        } else {
          const converted = convertNumber(f64(record.data, 6), xf, globals)
          setFormulaCell(sheet, row, col, converted.value, xf, globals, readStyles, formula)
        }
        break
      }
      case BIFF_STRING: {
        if (!pendingFormulaString) break
        const { value } = readBiff8UnicodeString(record.data, 0, globals.codePage)
        setFormulaCell(
          sheet,
          pendingFormulaString.row,
          pendingFormulaString.col,
          value,
          pendingFormulaString.xfIndex,
          globals,
          readStyles,
          pendingFormulaString.formula,
        )
        pendingFormulaString = undefined
        break
      }
      case BIFF_BLANK: {
        if (record.data.length < 6) break
        const row = u16(record.data, 0)
        const col = u16(record.data, 2)
        if (maxRows > 0 && row >= maxRows) break
        if (!inRange(row, col, range)) break
        setBlank(sheet, row, col, u16(record.data, 4), globals, readStyles)
        break
      }
      case BIFF_MULBLANK: {
        if (record.data.length < 6) break
        const row = u16(record.data, 0)
        if (maxRows > 0 && row >= maxRows) break
        const firstCol = u16(record.data, 2)
        const lastCol = u16(record.data, record.data.length - 2)
        let pos = 4
        for (let col = firstCol; col <= lastCol && pos + 2 <= record.data.length - 2; col++) {
          if (inRange(row, col, range)) setBlank(sheet, row, col, u16(record.data, pos), globals, readStyles)
          pos += 2
        }
        break
      }
      case BIFF_MERGECELLS: {
        const merges = parseMergeCells(record.data)
        if (merges.length > 0) sheet.merges = [...(sheet.merges ?? []), ...merges]
        break
      }
      case BIFF_HLINK: {
        const link = parseHLink(record.data)
        if (link) hyperlinks.push(link)
        break
      }
      case BIFF_ROW: {
        if (record.data.length < 16) break
        const row = u16(record.data, 0)
        const height = u16(record.data, 6)
        const flags = u32(record.data, 12)
        const hidden = (flags & 0x20) !== 0
        const outlineLevel = flags & 0x07
        const collapsed = (flags & 0x10) !== 0
        if (height || hidden || outlineLevel || collapsed) {
          if (!sheet.rowDefs) sheet.rowDefs = new Map()
          const def = sheet.rowDefs.get(row) ?? {}
          if (height) def.height = height / 20
          if (hidden) def.hidden = true
          if (outlineLevel) def.outlineLevel = outlineLevel
          if (collapsed) def.collapsed = true
          sheet.rowDefs.set(row, def)
        }
        break
      }
      case BIFF_COLINFO: {
        if (record.data.length < 12) break
        const first = u16(record.data, 0)
        const last = u16(record.data, 2)
        const width = u16(record.data, 4) / 256
        const flags = u16(record.data, 8)
        const hidden = (flags & 0x0001) !== 0
        const outlineLevel = (flags >> 8) & 0x07
        const collapsed = (flags & 0x1000) !== 0
        if (!sheet.columns) sheet.columns = []
        for (let col = first; col <= last; col++) {
          while (sheet.columns.length <= col) sheet.columns.push({})
          const def = sheet.columns[col] ?? {}
          if (width) def.width = width
          if (hidden) def.hidden = true
          if (outlineLevel) def.outlineLevel = outlineLevel
          if (collapsed) def.collapsed = true
          sheet.columns[col] = def
        }
        break
      }
    }
  }

  for (const link of hyperlinks) {
    const row = link.range.startRow
    const col = link.range.startCol
    if (maxRows > 0 && row >= maxRows) continue
    if (!inRange(row, col, range)) continue
    while (sheet.rows.length <= row) sheet.rows.push([])
    const arr = sheet.rows[row]!
    while (arr.length <= col) arr.push(null)
    if (!sheet.cells) sheet.cells = new Map()
    const key = `${row},${col}`
    const existing = sheet.cells.get(key) ?? { value: arr[col] ?? null, type: typeof arr[col] === "string" ? "string" : "empty" as CellType }
    existing.hyperlink = {
      target: link.target,
      ...(link.location ? { location: link.location } : {}),
      ...(link.tooltip ? { tooltip: link.tooltip } : {}),
    }
    sheet.cells.set(key, existing)
  }

  return sheet
}

/** Parse a BIFF Workbook stream into the hucre Workbook model. */
export function parseBiffWorkbook(workbookStream: Uint8Array, options?: ReadOptions): Workbook {
  const records = readRecords(workbookStream)
  const globals = parseGlobals(records)

  if (globals.sheets.length === 0) {
    throw new ParseError("Invalid XLS: workbook contains no BoundSheet records")
  }

  const sheets = globals.sheets
    .filter((s) => s.type === 0 && shouldReadSheet(s, options?.sheets))
    .map((sheet) => parseWorksheet(workbookStream, sheet, globals, options))

  const workbook: Workbook = {
    sheets,
    dateSystem: globals.dateSystem,
  }
  if (globals.namedRanges.length > 0) workbook.namedRanges = globals.namedRanges
  return workbook
}
