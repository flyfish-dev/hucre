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
import { readXlsx } from "./xlsx/reader"
import { writeXlsx } from "./xlsx/writer"
import { readOds } from "./ods/reader"
import { writeOds } from "./ods/writer"
import { EncryptedFileError, UnsupportedFormatError } from "./errors"
import { isOle2Container, readInputToUint8Array } from "./_input"

function detectFormat(data: Uint8Array): "xlsx" | "ods" {
  if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
    throw new UnsupportedFormatError("unknown (not a ZIP archive)")
  }
  if (data.length < 30) throw new UnsupportedFormatError("unknown (ZIP too short)")
  const filenameLen = data[26]! | (data[27]! << 8)
  if (data.length < 30 + filenameLen) throw new UnsupportedFormatError("unknown (ZIP truncated)")
  const decoder = new TextDecoder("utf-8")
  const firstName = decoder.decode(data.subarray(30, 30 + filenameLen))
  if (firstName === "mimetype") return "ods"
  return "xlsx"
}

export async function read(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = await readInputToUint8Array(input)
  if (isOle2Container(data)) throw new EncryptedFileError()
  const format = detectFormat(data)
  if (format === "ods") return readOds(data, options)
  return readXlsx(data, options)
}

export async function write(options: WriteOptions & { format?: "xlsx" | "ods" }): Promise<WriteOutput> {
  const format = options.format ?? "xlsx"
  if (format === "ods") return writeOds(options)
  return writeXlsx(options)
}

export async function readObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: ReadInput,
  options?: ReadOptions,
): Promise<T[]> {
  const workbook = await read(input, options)
  if (workbook.sheets.length === 0) return []
  const sheet = workbook.sheets[0]!
  const rows = sheet.rows
  if (rows.length === 0) return []
  const headers = rows[0]!.map((h) => (h === null || h === undefined ? "" : String(h).trim()))
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

export interface WriteObjectsTableOption {
  name: string
  style?: string
  showTotalRow?: boolean
  showAutoFilter?: boolean
  showRowStripes?: boolean
  totals?: Record<string, "sum" | "average" | "count" | "min" | "max" | "countNums" | "stdDev" | "var">
}

export async function writeObjects(
  data: Array<Record<string, CellValue>>,
  options?: { sheetName?: string; format?: "xlsx" | "ods"; table?: WriteObjectsTableOption },
): Promise<WriteOutput> {
  const sheetName = options?.sheetName ?? "Sheet1"
  const format = options?.format ?? "xlsx"
  if (data.length === 0) return write({ sheets: [{ name: sheetName, rows: [] }], format })
  const keys = Object.keys(data[0]!)
  const rows: CellValue[][] = [keys]
  for (const item of data) rows.push(keys.map((key) => item[key] ?? null))
  let tables: TableDefinition[] | undefined
  if (options?.table) {
    const t = options.table
    const endCol = String.fromCharCode(65 + keys.length - 1)
    const tableColumns: TableColumn[] = keys.map((key) => ({ name: key }))
    tables = [{ name: t.name, displayName: t.name, range: `A1:${endCol}${data.length + 1}`, columns: tableColumns }]
  }
  return write({ sheets: [{ name: sheetName, rows, tables }], format })
}
