// ── XLS CFB Property Set Reader ──────────────────────────────────────
// Parses the SummaryInformation and DocumentSummaryInformation streams in
// legacy OLE2/CFB XLS files. These streams are defined by the OLE property
// set format and surface the same workbook metadata that XLSX/XLSB expose
// through docProps/*.xml.

import type { WorkbookProperties } from "../_types"
import { CfbReader } from "./cfb"

const SUMMARY_INFORMATION = "\u0005SummaryInformation"
const DOCUMENT_SUMMARY_INFORMATION = "\u0005DocumentSummaryInformation"

const VT_I2 = 0x0002
const VT_I4 = 0x0003
const VT_R4 = 0x0004
const VT_R8 = 0x0005
const VT_DATE = 0x0007
const VT_BSTR = 0x0008
const VT_BOOL = 0x000b
const VT_LPSTR = 0x001e
const VT_LPWSTR = 0x001f
const VT_FILETIME = 0x0040
const VT_I8 = 0x0014
const VT_UI8 = 0x0015

// SummaryInformation property ids (MS-OLEPS / OLE property sets).
const PID_CODEPAGE = 0x0001
const PIDSI_TITLE = 0x0002
const PIDSI_SUBJECT = 0x0003
const PIDSI_AUTHOR = 0x0004
const PIDSI_KEYWORDS = 0x0005
const PIDSI_COMMENTS = 0x0006
const PIDSI_LASTAUTHOR = 0x0008
const PIDSI_CREATE_DTM = 0x000c
const PIDSI_LASTSAVE_DTM = 0x000d

// DocumentSummaryInformation property ids.
const PIDDSI_CATEGORY = 0x0002
const PIDDSI_MANAGER = 0x000e
const PIDDSI_COMPANY = 0x000f

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function u16(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint16(offset, true)
}

function u32(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint32(offset, true)
}

function i32(bytes: Uint8Array, offset: number): number {
  return view(bytes).getInt32(offset, true)
}

function f64(bytes: Uint8Array, offset: number): number {
  return view(bytes).getFloat64(offset, true)
}

function decodeUtf16Le(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-16le").decode(bytes)
  } catch {
    let out = ""
    for (let i = 0; i + 1 < bytes.length; i += 2) out += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8))
    return out
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

function decodeCodePage(bytes: Uint8Array, codePage: number): string {
  try {
    return new TextDecoder(codePageToEncoding(codePage)).decode(bytes)
  } catch {
    let out = ""
    for (const b of bytes) out += String.fromCharCode(b)
    return out
  }
}

function trimNulls(value: string): string {
  return value.replace(/\u0000+$/g, "")
}

function fileTimeToDate(low: number, high: number): Date | undefined {
  const raw = (BigInt(high) << 32n) | BigInt(low)
  if (raw === 0n) return undefined
  // FILETIME is 100-ns ticks since 1601-01-01 UTC. Unix epoch delta is
  // 116444736000000000 ticks.
  const unixMs = Number((raw - 116444736000000000n) / 10000n)
  if (!Number.isFinite(unixMs)) return undefined
  return new Date(unixMs)
}

function oleDateToDate(value: number): Date | undefined {
  if (!Number.isFinite(value)) return undefined
  const ms = Math.round((value - 25569) * 86_400_000)
  return new Date(ms)
}

function readStringValue(data: Uint8Array, offset: number, codePage: number, wide: boolean): string | undefined {
  if (offset + 4 > data.length) return undefined
  const length = u32(data, offset)
  const start = offset + 4
  if (wide) {
    const byteLen = length * 2
    if (start + byteLen > data.length) return undefined
    return trimNulls(decodeUtf16Le(data.subarray(start, start + byteLen)))
  }
  if (start + length > data.length) return undefined
  return trimNulls(decodeCodePage(data.subarray(start, start + length), codePage))
}

function readTypedValue(data: Uint8Array, offset: number, codePage: number): unknown {
  if (offset + 4 > data.length) return undefined
  const vt = u32(data, offset) & 0xffff
  const valueOffset = offset + 4

  switch (vt) {
    case VT_I2:
      return valueOffset + 2 <= data.length ? view(data).getInt16(valueOffset, true) : undefined
    case VT_I4:
      return valueOffset + 4 <= data.length ? i32(data, valueOffset) : undefined
    case VT_I8:
    case VT_UI8:
      if (valueOffset + 8 > data.length) return undefined
      return Number(view(data).getBigInt64(valueOffset, true))
    case VT_R4:
      return valueOffset + 4 <= data.length ? view(data).getFloat32(valueOffset, true) : undefined
    case VT_R8:
      return valueOffset + 8 <= data.length ? f64(data, valueOffset) : undefined
    case VT_DATE:
      return valueOffset + 8 <= data.length ? oleDateToDate(f64(data, valueOffset)) : undefined
    case VT_BOOL:
      return valueOffset + 2 <= data.length ? view(data).getInt16(valueOffset, true) !== 0 : undefined
    case VT_BSTR:
    case VT_LPSTR:
      return readStringValue(data, valueOffset, codePage, false)
    case VT_LPWSTR:
      return readStringValue(data, valueOffset, codePage, true)
    case VT_FILETIME:
      if (valueOffset + 8 > data.length) return undefined
      return fileTimeToDate(u32(data, valueOffset), u32(data, valueOffset + 4))
    default:
      return undefined
  }
}

function parsePropertySection(data: Uint8Array, sectionOffset: number): Map<number, unknown> {
  const props = new Map<number, unknown>()
  if (sectionOffset < 0 || sectionOffset + 8 > data.length) return props

  const propCount = u32(data, sectionOffset + 4)
  let codePage = 1252
  const entries: Array<{ id: number; offset: number }> = []
  let pos = sectionOffset + 8
  for (let i = 0; i < propCount && pos + 8 <= data.length; i++) {
    entries.push({ id: u32(data, pos), offset: u32(data, pos + 4) })
    pos += 8
  }

  // Code page affects LPSTR decoding; read it first when present.
  const codePageEntry = entries.find((e) => e.id === PID_CODEPAGE)
  if (codePageEntry && sectionOffset + codePageEntry.offset + 6 <= data.length) {
    const value = readTypedValue(data, sectionOffset + codePageEntry.offset, codePage)
    if (typeof value === "number") codePage = value
  }

  for (const entry of entries) {
    const absolute = sectionOffset + entry.offset
    if (absolute < 0 || absolute + 4 > data.length) continue
    const value = readTypedValue(data, absolute, codePage)
    if (value !== undefined) props.set(entry.id, value)
  }

  return props
}

function parsePropertySet(data: Uint8Array): Map<number, unknown> {
  const result = new Map<number, unknown>()
  if (data.length < 48 || u16(data, 0) !== 0xfffe) return result
  const sectionCount = u32(data, 24)
  let pos = 28
  for (let i = 0; i < sectionCount && pos + 20 <= data.length; i++) {
    const sectionOffset = u32(data, pos + 16)
    for (const [key, value] of parsePropertySection(data, sectionOffset)) result.set(key, value)
    pos += 20
  }
  return result
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 ? value : undefined
}

function asDate(value: unknown): Date | undefined {
  return value instanceof Date && !Number.isNaN(value.getTime()) ? value : undefined
}

/** Extract workbook metadata from XLS CFB property streams, if present. */
export function parseXlsProperties(cfb: CfbReader): WorkbookProperties | undefined {
  const props: WorkbookProperties = {}

  const summary = cfb.getStream(SUMMARY_INFORMATION)
  if (summary) {
    const s = parsePropertySet(summary)
    const title = asString(s.get(PIDSI_TITLE))
    const subject = asString(s.get(PIDSI_SUBJECT))
    const creator = asString(s.get(PIDSI_AUTHOR))
    const keywords = asString(s.get(PIDSI_KEYWORDS))
    const description = asString(s.get(PIDSI_COMMENTS))
    const lastModifiedBy = asString(s.get(PIDSI_LASTAUTHOR))
    const created = asDate(s.get(PIDSI_CREATE_DTM))
    const modified = asDate(s.get(PIDSI_LASTSAVE_DTM))
    if (title) props.title = title
    if (subject) props.subject = subject
    if (creator) props.creator = creator
    if (keywords) props.keywords = keywords
    if (description) props.description = description
    if (lastModifiedBy) props.lastModifiedBy = lastModifiedBy
    if (created) props.created = created
    if (modified) props.modified = modified
  }

  const docSummary = cfb.getStream(DOCUMENT_SUMMARY_INFORMATION)
  if (docSummary) {
    const d = parsePropertySet(docSummary)
    const category = asString(d.get(PIDDSI_CATEGORY))
    const manager = asString(d.get(PIDDSI_MANAGER))
    const company = asString(d.get(PIDDSI_COMPANY))
    if (category) props.category = category
    if (manager) props.manager = manager
    if (company) props.company = company
  }

  return Object.keys(props).length > 0 ? props : undefined
}
