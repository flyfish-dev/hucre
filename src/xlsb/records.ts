// ── XLSB BIFF12 Record Utilities ────────────────────────────────────
// XLSB worksheet/workbook parts are streams of BIFF12 records. Record ids
// and payload lengths are variable-length integers followed by a payload.

import { ParseError } from "../errors"

export interface XlsbRecord {
  type: number
  offset: number
  data: Uint8Array
}

export const BRT_ROW_HDR = 0x0000
export const BRT_CELL_BLANK = 0x0001
export const BRT_CELL_RK = 0x0002
export const BRT_CELL_ERROR = 0x0003
export const BRT_CELL_BOOL = 0x0004
export const BRT_CELL_REAL = 0x0005
export const BRT_CELL_ST = 0x0006
export const BRT_CELL_ISST = 0x0007
export const BRT_FMLA_STRING = 0x0008
export const BRT_FMLA_NUM = 0x0009
export const BRT_FMLA_BOOL = 0x000a
export const BRT_FMLA_ERROR = 0x000b
export const BRT_SHORT_BLANK = 0x000c
export const BRT_SHORT_RK = 0x000d
export const BRT_SHORT_ERROR = 0x000e
export const BRT_SHORT_BOOL = 0x000f
export const BRT_SHORT_REAL = 0x0010
export const BRT_SHORT_ST = 0x0011
export const BRT_SHORT_ISST = 0x0012
export const BRT_SST_ITEM = 0x0013
export const BRT_MERGE_CELL = 0x00b0
export const BRT_WS_DIM = 0x0094
export const BRT_BUNDLE_SH = 0x009c
export const BRT_WB_PROP = 0x0099
export const BRT_FMT = 0x002c
export const BRT_XF = 0x002f
export const BRT_HLINK = 0x01ee
export const BRT_BEGIN_CELL_XFS = 0x0269
export const BRT_END_CELL_XFS = 0x026a

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

export function u16(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint16(offset, true)
}

export function u32(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint32(offset, true)
}

export function i32(bytes: Uint8Array, offset: number): number {
  return view(bytes).getInt32(offset, true)
}

export function f64(bytes: Uint8Array, offset: number): number {
  return view(bytes).getFloat64(offset, true)
}

function readVarInt(bytes: Uint8Array, offset: number): { value: number; offset: number } {
  let pos = offset
  let value = 0
  let shift = 0
  for (let i = 0; i < 5; i++) {
    if (pos >= bytes.length) throw new ParseError("Invalid XLSB record header: truncated varint")
    const b = bytes[pos++]!
    value |= (b & 0x7f) << shift
    if ((b & 0x80) === 0) return { value, offset: pos }
    shift += 7
  }
  throw new ParseError("Invalid XLSB record header: varint too long")
}

export function readXlsbRecords(bytes: Uint8Array): XlsbRecord[] {
  const records: XlsbRecord[] = []
  let pos = 0
  while (pos < bytes.length) {
    const recordOffset = pos
    const typeInfo = readVarInt(bytes, pos)
    const lenInfo = readVarInt(bytes, typeInfo.offset)
    pos = lenInfo.offset
    const end = pos + lenInfo.value
    if (end > bytes.length) {
      throw new ParseError(
        `Invalid XLSB record 0x${typeInfo.value.toString(16)}: payload extends past part end`,
      )
    }
    records.push({ type: typeInfo.value, offset: recordOffset, data: bytes.subarray(pos, end) })
    pos = end
  }
  return records
}

export function decodeUtf16Le(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-16le").decode(bytes)
  } catch {
    let s = ""
    for (let i = 0; i + 1 < bytes.length; i += 2) s += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8))
    return s
  }
}

/**
 * XLSB XLWideString: 4-byte character count followed by UTF-16LE chars.
 */
export function readXlsbWideString(data: Uint8Array, offset: number): { value: string; offset: number } {
  if (offset + 4 > data.length) return { value: "", offset: data.length }
  const cch = u32(data, offset)
  const start = offset + 4
  const byteLen = cch * 2
  const end = Math.min(start + byteLen, data.length)
  return { value: decodeUtf16Le(data.subarray(start, end)), offset: end }
}

/**
 * Common XLSB cell header: column (4), style index (3), flags (1).
 */
export function readCellHeader(data: Uint8Array): { col: number; style: number; offset: number } | null {
  if (data.length < 8) return null
  const col = u32(data, 0)
  const style = (data[4] ?? 0) | ((data[5] ?? 0) << 8) | ((data[6] ?? 0) << 16)
  return { col, style, offset: 8 }
}

/**
 * Short XLSB cell records carry the current BrtRowHdr row and a compact
 * column/value payload without an explicit cell XF. Treat them as style 0
 * while preserving the cached value. This covers the BrtShort* family in
 * MS-XLSB and is harmless for regular BrtCell* records, which continue to
 * use {@link readCellHeader}.
 */
export function readShortCellHeader(data: Uint8Array): { col: number; style: number; offset: number } | null {
  if (data.length < 4) return null
  return { col: u32(data, 0), style: 0, offset: 4 }
}

export function decodeRk(raw: number): number {
  let value: number
  if ((raw & 0x02) !== 0) {
    value = raw >> 2
  } else {
    const buf = new ArrayBuffer(8)
    const dv = new DataView(buf)
    dv.setUint32(0, 0, true)
    dv.setUint32(4, raw & 0xfffffffc, true)
    value = dv.getFloat64(0, true)
  }
  if ((raw & 0x01) !== 0) value /= 100
  return value
}

export function decodeError(code: number): string {
  return ERROR_TEXT[code] ?? `#ERR${code}`
}
