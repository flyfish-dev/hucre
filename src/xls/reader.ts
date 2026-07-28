// ── XLS Reader ───────────────────────────────────────────────────────
// Reads legacy Excel binary workbooks (.xls). XLS files are normally
// Compound File Binary (OLE2/CFB) containers whose Workbook/Book stream is
// encoded as BIFF records. Very old BIFF workbooks can be a raw BIFF stream;
// direct readXls() accepts those too.

import type { Workbook, ReadInput, ReadOptions } from "../_types"
import { ParseError } from "../errors"
import { readInputToUint8Array } from "../_input"
import { decryptOfficeEncryptedPackage, isOfficeEncryptedPackage } from "../crypto/office-crypto"
import { parseBiffWorkbook } from "./biff"
import { CfbReader } from "./cfb"
import { parseXlsProperties } from "./properties"

const CFB_MAGIC = Object.freeze([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const)
const BIFF_BOF = 0x0809
const BIFF2_BOF = 0x0009
const BIFF3_BOF = 0x0209
const BIFF4_BOF = 0x0409

interface BinaryWorkbookPart {
  path: string
  kind: "vba" | "drawing" | "chart" | "comment" | "pivot" | "table" | "metadata" | "unknown"
  data: Uint8Array
}

function u16(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint16(offset, true)
}

function isCfb(data: Uint8Array): boolean {
  if (data.length < CFB_MAGIC.length) return false
  for (let i = 0; i < CFB_MAGIC.length; i++) if (data[i] !== CFB_MAGIC[i]) return false
  return true
}

function isRawBiff(data: Uint8Array): boolean {
  if (data.length < 4) return false
  const sid = u16(data, 0)
  return sid === BIFF_BOF || sid === BIFF2_BOF || sid === BIFF3_BOF || sid === BIFF4_BOF
}

/**
 * Read a legacy Excel XLS file and return a Workbook.
 *
 * The implementation supports the OLE2/CFB-hosted BIFF5-BIFF8 workbook
 * streams used by Excel 5.0 through Excel 2003 and raw BIFF streams used by
 * older exports. Password-protected workbooks are detected through FilePass
 * / encrypted-package markers and surfaced as {@link EncryptedFileError}.
 */
export async function readXls(
  input: ReadInput,
  options?: ReadOptions & { password?: string },
): Promise<Workbook> {
  const data = await readInputToUint8Array(input)

  if (!isCfb(data)) {
    if (isRawBiff(data)) return parseBiffSafely(data, options)
    throw new ParseError("Invalid XLS: missing OLE2/CFB header or BIFF BOF record")
  }

  if (isOfficeEncryptedPackage(data)) {
    const decrypted = await decryptOfficeEncryptedPackage(data, options?.password, "xls")
    if (isRawBiff(decrypted)) return parseBiffSafely(decrypted, options)
    if (isCfb(decrypted)) return readXls(decrypted, options)
    throw new ParseError(
      "Encrypted package decrypted successfully, but it does not contain a legacy XLS BIFF workbook",
    )
  }

  const cfb = new CfbReader(data)

  const workbookStream = cfb.getStream("Workbook") ?? cfb.getStream("Book")
  if (!workbookStream) {
    throw new ParseError("Invalid XLS: missing Workbook/Book stream")
  }

  const workbook = parseBiffSafely(workbookStream, options)
  const properties = parseXlsProperties(cfb)
  if (properties) workbook.properties = properties

  const parts = collectBinaryParts(cfb)
  if (parts.length > 0) {
    const target = workbook as Workbook & {
      binaryParts?: BinaryWorkbookPart[]
      vbaProject?: { parts: BinaryWorkbookPart[] }
    }
    target.binaryParts = parts
    const vbaParts = parts.filter((p) => p.kind === "vba")
    if (vbaParts.length > 0) target.vbaProject = { parts: vbaParts }
  }

  return workbook
}

function parseBiffSafely(
  workbookStream: Uint8Array,
  options?: ReadOptions & { password?: string },
): Workbook {
  try {
    return parseBiffWorkbook(workbookStream, options)
  } catch (err) {
    if (err instanceof ParseError) throw err
    throw new ParseError("Failed to parse XLS workbook (malformed or truncated)", undefined, {
      cause: err,
    })
  }
}

function collectBinaryParts(cfb: CfbReader): BinaryWorkbookPart[] {
  const parts: BinaryWorkbookPart[] = []
  for (const entry of cfb.listStreams()) {
    const normalized = entry.name.replace(/^\u0005/, "")
    if (/^(Workbook|Book|SummaryInformation|DocumentSummaryInformation)$/i.test(normalized))
      continue
    const data = cfb.getStream(entry.name)
    if (!data || data.length === 0) continue
    parts.push({ path: entry.name, kind: classifyCfbPart(entry.name), data })
  }
  return parts
}

function classifyCfbPart(name: string): BinaryWorkbookPart["kind"] {
  const n = name.toLowerCase()
  if (n.includes("vba") || n.includes("_vba_project")) return "vba"
  if (n.includes("drawing") || n.includes("mso") || n.includes("escher")) return "drawing"
  if (n.includes("chart")) return "chart"
  if (n.includes("comment") || n.includes("note")) return "comment"
  if (n.includes("pivot")) return "pivot"
  if (n.includes("table")) return "table"
  if (n.includes("metadata")) return "metadata"
  return "unknown"
}
