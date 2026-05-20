// ── High-Level API ──────────────────────────────────────────────────
export { read, write, readObjects, writeObjects } from "./defter"
export type { WriteObjectsTableOption } from "./defter"
export type * from "./_types"
export {
  DefterError,
  EncryptedFileError,
  ParseError,
  UnsupportedFormatError,
  ValidationError,
  XmlError,
  ZipError,
} from "./errors"

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

// ── Office encryption helpers ──────────────────────────────────────
export {
  decryptOfficeEncryptedPackage,
  encryptOfficeAgilePackage,
  encryptOfficeAgilePackageParts,
  isOfficeEncryptedPackage,
} from "./crypto/office-crypto"
export type {
  AgileEncryptionOptions,
  EncryptedOfficePackageParts,
  OfficeCryptoOptions,
} from "./crypto/office-crypto"

// ── ODS ────────────────────────────────────────────────────────────
export { readOds } from "./ods/reader"
export { writeOds } from "./ods/writer"
export { readOdsObjects, writeOdsObjects } from "./ods/objects"
export type { OdsObjectsReadOptions, OdsObjectsResult, OdsObjectsWriteOptions } from "./ods/objects"

// ── CSV ────────────────────────────────────────────────────────────
export {
  detectDelimiter,
  fetchCsv,
  formatCsvValue,
  parseCsv,
  parseCsvObjects,
  stripBom,
  writeCsv,
  writeCsvObjects,
} from "./csv/index"
export { CsvStreamWriter, streamCsvRows } from "./csv/stream"

// ── JSON ───────────────────────────────────────────────────────────
export {
  collectHeaders,
  flattenValue,
  NdjsonStreamWriter,
  parseJson,
  parseNdjson,
  parseValue,
  readNdjsonStream,
  workbookToJson,
  writeJson,
  writeNdjson,
} from "./json"
export type {
  FlattenOptions,
  JsonReadOptions,
  JsonReadResult,
  JsonWriteOptions,
  NdjsonReadOptions,
  NdjsonStreamReadOptions,
  WorkbookToJsonOptions,
} from "./json"

// ── Schema, dates, and number formats ──────────────────────────────
export { validateWithSchema } from "./_schema"
export {
  dateToSerial,
  formatDate,
  isDateFormat,
  parseDate,
  serialToDate,
  serialToTime,
  timeToSerial,
} from "./_date"
export { formatValue } from "./_format"
export type { FormatOptions } from "./_format"

// ── Sheet and cell utilities ───────────────────────────────────────
export {
  cloneSheet,
  copyRange,
  copySheetToWorkbook,
  deleteColumns,
  deleteRows,
  findCells,
  groupRows,
  hideColumns,
  hideRows,
  insertColumns,
  insertRows,
  moveRows,
  moveSheet,
  removeSheet,
  replaceCells,
  sortRows,
} from "./sheet-ops"
export { sheetToArrays, sheetToObjects } from "./sheet-utils"
export {
  a1ToR1C1,
  cellRef,
  colToLetter,
  isInRange,
  letterToCol,
  parseCellRef,
  parseRange,
  r1c1ToA1,
  rangeRef,
} from "./cell-utils"

// ── Export helpers ─────────────────────────────────────────────────
export { fromHtml, toHtml, toJson, toMarkdown } from "./export/index"
export type { HtmlExportOptions, JsonExportOptions, MarkdownExportOptions } from "./export/index"

// ── Worker and image helpers ───────────────────────────────────────
export { deserializeWorkbook, serializeWorkbook, WORKER_SAFE_FUNCTIONS } from "./worker"
export type {
  SerializedCell,
  SerializedCellValue,
  SerializedSheet,
  SerializedSheetImage,
  SerializedWorkbook,
  SerializedWorkbookProperties,
} from "./worker"
export { imageFromBase64 } from "./image"
