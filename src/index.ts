// ── High-Level API ──────────────────────────────────────────────────
export { read, write, readObjects, writeObjects } from "./defter"
export type { WriteObjectsTableOption } from "./defter"

// ── XLSX ────────────────────────────────────────────────────────────
export { readXlsx } from "./xlsx/reader"
export { writeXlsx } from "./xlsx/writer"
export { openXlsx, saveXlsx } from "./xlsx/roundtrip"
export type { RoundtripWorkbook } from "./xlsx/roundtrip"
export { hashSheetPassword } from "./xlsx/password"
export { calculateColumnWidth, measureValueWidth, calculateRowHeight } from "./xlsx/auto-size"
export { parseThemeColors, resolveThemeColor } from "./xlsx/theme"
export { streamXlsxRows } from "./xlsx/stream-reader"
export type { StreamRow } from "./xlsx/stream-reader"
export { XlsxStreamWriter } from "./xlsx/stream-writer"
export type { StreamWriterOptions } from "./xlsx/stream-writer"
export { readXlsxObjects, writeXlsxObjects } from "./xlsx/objects"
export type {
  XlsxObjectsReadOptions,
  XlsxObjectsResult,
  XlsxObjectsWriteOptions,
} from "./xlsx/objects"

// ── XLS ─────────────────────────────────────────────────────────────
export { readXls } from "./xls/reader"

// ── XLSB ────────────────────────────────────────────────────────────
export { readXlsb } from "./xlsb/reader"

// ── ODS ────────────────────────────────────────────────────────────
export { readOds } from "./ods/reader"
