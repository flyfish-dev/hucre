// ── ReadInput Normalization ────────────────────────────────────────────

import type { ReadInput } from "./_types"
import { EncryptedFileError, ParseError } from "./errors"
import type { WorkbookFormat } from "./errors"

const OLE2_MAGIC = Object.freeze([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const)

/** Whether `data` starts with the OLE2 / CFB compound-document magic bytes. */
export function isOle2Container(data: Uint8Array): boolean {
  if (data.length < OLE2_MAGIC.length) return false
  for (let i = 0; i < OLE2_MAGIC.length; i++) {
    if (data[i] !== OLE2_MAGIC[i]) return false
  }
  return true
}

/** Throw {@link EncryptedFileError} when `data` is an encrypted OOXML/ODF CFB envelope. */
export function assertNotEncrypted(data: Uint8Array, format: WorkbookFormat): void {
  if (isOle2Container(data)) {
    throw new EncryptedFileError(format)
  }
}

/** Drain a ReadableStream of byte chunks into one Uint8Array. */
export async function bufferReadableStream(
  stream: ReadableStream<Uint8Array>,
): Promise<Uint8Array> {
  const reader = stream.getReader()
  const chunks: Uint8Array[] = []
  let totalLen = 0

  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    if (value) {
      chunks.push(value)
      totalLen += value.length
    }
  }

  if (chunks.length === 0) return new Uint8Array(0)
  if (chunks.length === 1) return chunks[0]!

  const result = new Uint8Array(totalLen)
  let offset = 0
  for (const chunk of chunks) {
    result.set(chunk, offset)
    offset += chunk.length
  }
  return result
}

function isReadableStream(value: unknown): value is ReadableStream<Uint8Array> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as ReadableStream<Uint8Array>).getReader === "function"
  )
}

/** Normalize ReadInput into Uint8Array. */
export async function readInputToUint8Array(input: ReadInput): Promise<Uint8Array> {
  if (input instanceof Uint8Array) return input
  if (input instanceof ArrayBuffer) return new Uint8Array(input)
  if (isReadableStream(input)) return bufferReadableStream(input)
  throw new ParseError(
    "Unsupported input type. Expected Uint8Array, ArrayBuffer, or ReadableStream<Uint8Array>.",
  )
}
