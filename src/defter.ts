// ── Ergonomic API ───────────────────────────────────────────────────
// Unified high-level functions that wrap format-specific readers/writers.
// Auto-detects format from content (magic bytes / package metadata) for
// reading, and dispatches to the correct writer based on the format option.

import type {
  Workbook,
  ReadOptions,
  WriteOptions,
  WriteOutput,
  CellValue,
  ReadInput,
  TableDefinition,
  TableColumn,
} from "./_types"
import { readXls } from "./xls/reader"
import { readXlsx } from "./xlsx/reader"
import { readXlsb } from "./xlsb/reader"
import { writeXlsx } from "./xlsx/writer"
import { readOds } from "./ods/reader"
import { writeOds } from "./ods/writer"
import { EncryptedFileError, UnsupportedFormatError } from "./errors"
import { isOle2Container, readInputToUint8Array } from "./_input"
import { decryptOfficeEncryptedPackage, isOfficeEncryptedPackage } from "./crypto/office-crypto"
import { ZipReader } from "./zip/reader"
import { parseXml } from "./xml/parser"
import type { XmlElement } from "./xml/parser"
import { decryptAgile } from "./xlsx/crypto/agile"

// ── Format Detection ────────────────────────────────────────────────

function isZip(data: Uint8Array): boolean {
  return data.length >= 4 && data[0] === 0x50 && data[1] === 0x4b
}

function u16(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint16(offset, true)
}

function isRawBiff(data: Uint8Array): boolean {
  if (data.length < 4) return false
  const sid = u16(data, 0)
  return sid === 0x0809 || sid === 0x0009 || sid === 0x0209 || sid === 0x0409
}

/**
 * Detect whether a ZIP archive is XLSX, XLSB, or ODS by inspecting the
 * package metadata rather than the extension.
 */
async function detectZipFormat(data: Uint8Array): Promise<"xlsx" | "xlsb" | "ods"> {
  if (!isZip(data)) {
    throw new UnsupportedFormatError("unknown (not a ZIP archive)")
  }

  if (data.length < 30) {
    throw new UnsupportedFormatError("unknown (ZIP too short)")
  }

  const decoder = new TextDecoder("utf-8")
  const filenameLen = data[26]! | (data[27]! << 8)
  if (data.length >= 30 + filenameLen) {
    const firstName = decoder.decode(data.subarray(30, 30 + filenameLen))
    if (firstName === "mimetype") {
      const extraLen = data[28]! | (data[29]! << 8)
      const dataOffset = 30 + filenameLen + extraLen
      const uncompSize = data[22]! | (data[23]! << 8) | (data[24]! << 16) | (data[25]! << 24)
      if (uncompSize > 0 && data.length >= dataOffset + uncompSize) {
        const mimeContent = decoder.decode(data.subarray(dataOffset, dataOffset + uncompSize))
        if (mimeContent.trim() === "application/vnd.oasis.opendocument.spreadsheet") return "ods"
      }
      return "ods"
    }
  }

  const zip = new ZipReader(data)
  if (zip.has("[Content_Types].xml")) {
    const contentTypes = decoder.decode(await zip.extract("[Content_Types].xml"))
    const workbookFormat = detectWorkbookFormatFromContentTypes(contentTypes)
    if (workbookFormat) return workbookFormat
  }

  if (zip.has("xl/workbook.xml")) return "xlsx"
  if (zip.has("xl/workbook.bin")) return "xlsb"
  if (zip.entries().some((entry) => /(^|\/)workbook\.bin$/i.test(entry))) return "xlsb"

  return "xlsx"
}

function detectWorkbookFormatFromContentTypes(xml: string): "xlsx" | "xlsb" | undefined {
  let root: XmlElement
  try {
    root = parseXml(xml)
  } catch {
    return undefined
  }

  const stack: XmlElement[] = [root]
  while (stack.length) {
    const el = stack.pop()!
    if (el.local === "Override") {
      const partName = normalizePackagePartName(el.attrs["PartName"])
      if (/(^|\/)workbook\.xml$/i.test(partName)) return "xlsx"
      if (/(^|\/)workbook\.bin$/i.test(partName)) return "xlsb"
    }
    for (const child of el.children) {
      if (typeof child !== "string") stack.push(child)
    }
  }

  return undefined
}

function normalizePackagePartName(partName = ""): string {
  return partName.replace(/^\/+/, "")
}

async function decryptDetectedOfficePackage(
  data: Uint8Array,
  password?: string,
): Promise<Uint8Array> {
  if (!password) {
    return decryptOfficeEncryptedPackage(data, password)
  }

  try {
    return await decryptAgile(data, password)
  } catch (agileError) {
    try {
      return await decryptOfficeEncryptedPackage(data, password)
    } catch {
      throw agileError
    }
  }
}

// ── Public API ──────────────────────────────────────────────────────

/**
 * Read any supported spreadsheet file. Auto-detects XLS, XLSX, XLSB, and ODS.
 * CSV uses parseCsv separately since it is string input.
 */
export async function read(
  input: ReadInput,
  options?: ReadOptions & { password?: string },
): Promise<Workbook> {
  let data = await readInputToUint8Array(input)

  if (isOle2Container(data)) {
    if (isOfficeEncryptedPackage(data)) {
      data = await decryptDetectedOfficePackage(data, options?.password)
    } else {
      // Preserve the historical byte-sniff behavior for short synthetic
      // encrypted-container probes, while allowing malformed real XLS files
      // to surface the parser's typed error instead of being mislabeled.
      if (data.length < 512) throw new EncryptedFileError()
      return readXls(data, options)
    }
  }

  if (isRawBiff(data)) {
    return readXls(data, options)
  }

  const format = await detectZipFormat(data)
  if (format === "ods") return readOds(data, options)
  if (format === "xlsb") return readXlsb(data, options)
  return readXlsx(data, options)
}

/** Write a workbook to the specified format. */
export async function write(
  options: WriteOptions & { format?: "xlsx" | "ods" },
): Promise<WriteOutput> {
  const format = options.format ?? "xlsx"
  if (format === "ods") return writeOds(options)
  return writeXlsx(options)
}

/** Quick helper: read a file and get the first sheet as array of objects. */
export async function readObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: ReadInput,
  options?: ReadOptions & { password?: string },
): Promise<T[]> {
  const workbook = await read(input, options)
  if (workbook.sheets.length === 0) return []

  const sheet = workbook.sheets[0]!
  const rows = sheet.rows
  if (rows.length === 0) return []

  const headers = rows[0]!.map((h) => (h === null || h === undefined ? "" : String(h).trim()))
  if (headers.length === 0) return []

  const data: T[] = []
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]!
    const obj: Record<string, CellValue> = {}
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j]!
      if (key === "") continue
      obj[key] = j < row.length ? (row[j] ?? null) : null
    }
    data.push(obj as T)
  }

  return data
}

/** Options for writeObjects table generation. */
export interface WriteObjectsTableOption {
  name: string
  style?: string
  showTotalRow?: boolean
  showAutoFilter?: boolean
  showRowStripes?: boolean
  totals?: Record<
    string,
    "sum" | "average" | "count" | "min" | "max" | "countNums" | "stdDev" | "var"
  >
}

/** Write an array of objects to a spreadsheet format. */
export async function writeObjects(
  data: Array<Record<string, CellValue>>,
  options?: {
    sheetName?: string
    format?: "xlsx" | "ods"
    table?: WriteObjectsTableOption
  },
): Promise<WriteOutput> {
  const sheetName = options?.sheetName ?? "Sheet1"
  const format = options?.format ?? "xlsx"

  if (data.length === 0) {
    return write({ sheets: [{ name: sheetName, rows: [] }], format })
  }

  const keys = Object.keys(data[0]!)
  const rows: CellValue[][] = [keys]
  for (const item of data) {
    rows.push(keys.map((key) => (item[key] === undefined ? null : item[key]!)))
  }

  let tables: TableDefinition[] | undefined
  if (options?.table) {
    const t = options.table
    const colCount = keys.length
    const rowCount = data.length + 1
    const endCol = colToLetterSimple(colCount - 1)
    const range = `A1:${endCol}${rowCount + (t.showTotalRow ? 1 : 0)}`
    const tableColumns: TableColumn[] = keys.map((key) => ({
      name: key,
      ...(t.totals?.[key] ? { totalFunction: t.totals[key] } : {}),
    }))
    tables = [
      {
        name: t.name,
        displayName: t.name,
        range,
        columns: tableColumns,
        style: t.style,
        showAutoFilter: t.showAutoFilter,
        showRowStripes: t.showRowStripes,
        showTotalRow: t.showTotalRow,
      },
    ]
  }

  return write({ sheets: [{ name: sheetName, rows, tables }], format })
}

function colToLetterSimple(col: number): string {
  let result = ""
  let n = col
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result
    n = Math.floor(n / 26) - 1
  }
  return result
}
