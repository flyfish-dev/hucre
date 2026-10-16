import type { SheetImage } from "../_types"
import { ParseError } from "../errors"
import type { BiffRecord } from "./biff"

const DRAWING_GROUP = 0x00eb
const DRAWING = 0x00ec
const CONTINUE = 0x003c
const MAX_DRAWING_BYTES = 64 * 1024 * 1024
const MAX_ART_RECORDS = 100_000
type ArtRecord = { type: number; instance: number; version: number; data: Uint8Array }
type Picture = Pick<SheetImage, "data" | "type">

/**
 * BIFF 的绘图跨 MSODRAWING / CONTINUE，不能把每条 BIFF 记录当完整图片。
 * 仅复制有上限的绘图部分，不复制 Workbook 全文；图片池由所有工作表共享。
 */
function drawingBytes(
  records: readonly BiffRecord[],
  start: number,
  end: number,
  id: number,
): Uint8Array {
  const parts: Uint8Array[] = []
  let length = 0
  for (let index = start; index < end; index++) {
    const record = records[index]!
    if (record.id !== id) continue
    let continued: boolean
    do {
      const part = records[index]!.data
      length += part.length
      if (length > MAX_DRAWING_BYTES)
        throw new ParseError("XLS drawing data exceeds the resource limit")
      parts.push(part)
      const next = records[index + 1]
      continued = !!next && index + 1 < end && next.id === CONTINUE
      if (continued) index++
    } while (continued)
  }
  const bytes = new Uint8Array(length)
  let offset = 0
  for (const part of parts) {
    bytes.set(part, offset)
    offset += part.length
  }
  return bytes
}

/** 所有长度均受当前容器边界约束；损坏绘图不能越界读取或无限递归。 */
function children(bytes: Uint8Array, budget: { remaining: number }): ArtRecord[] {
  const result: ArtRecord[] = []
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let offset = 0; offset + 8 <= bytes.length;) {
    if (--budget.remaining < 0)
      throw new ParseError("XLS drawing record count exceeds the resource limit")
    const options = view.getUint16(offset, true)
    const length = view.getUint32(offset + 4, true)
    if (length > bytes.length - offset - 8) break
    result.push({
      type: view.getUint16(offset + 2, true),
      instance: options >>> 4,
      version: options & 15,
      data: bytes.subarray(offset + 8, offset + 8 + length),
    })
    offset += 8 + length
  }
  return result
}

function picture(record: ArtRecord | undefined): Picture | null {
  if (!record || record.version !== 0) return null
  // [MS-ODRAW] OfficeArtBlipPNG / OfficeArtBlipJPEG：UID 数量由 recInstance 决定。
  const png = record.type === 0xf01e && [0x6e0, 0x6e1].includes(record.instance)
  const jpeg =
    [0xf01d, 0xf02a].includes(record.type) && [0x46a, 0x46b, 0x6e2, 0x6e3].includes(record.instance)
  if (!png && !jpeg) return null
  const offset = record.instance & 1 ? 33 : 17
  const data = record.data.subarray(offset)
  if (
    png &&
    (data.length < 24 ||
      ![137, 80, 78, 71, 13, 10, 26, 10].every((byte, index) => data[index] === byte))
  )
    return null
  if (jpeg && (data.length < 3 || data[0] !== 0xff || data[1] !== 0xd8 || data[2] !== 0xff))
    return null
  // 一次复制解除整份 BIFF 缓冲区的引用；多个图片实例继续共享这一份有效字节。
  return { type: png ? "png" : "jpeg", data: data.slice() }
}

/** 保留无效/暂不支持的池条目，避免后面的 pib 索引错配到另一张图片。 */
export function readXlsPicturePool(
  records: readonly BiffRecord[],
  globalsEnd: number,
): Array<Picture | null> {
  const bytes = drawingBytes(records, 0, globalsEnd, DRAWING_GROUP)
  const budget = { remaining: MAX_ART_RECORDS }
  const pool: Array<Picture | null> = []
  const visit = (data: Uint8Array, depth: number): void => {
    if (depth > 24) throw new ParseError("XLS drawing nesting exceeds the resource limit")
    for (const record of children(data, budget)) {
      if (record.type === 0xf007) {
        const body = record.data
        const embedded = body.length >= 36 ? body.subarray(36 + body[33]!) : new Uint8Array()
        pool.push(picture(children(embedded, budget)[0]))
      } else if (record.version === 15) visit(record.data, depth + 1)
    }
  }
  visit(bytes, 0)
  return pool
}

/** 读取当前工作表的真实 ClientAnchor；不按图片像素或默认单元格尺寸猜位置。 */
export function readXlsSheetImages(
  records: readonly BiffRecord[],
  start: number,
  pool: Array<Picture | null>,
): SheetImage[] {
  if (!pool.some(Boolean)) return []
  let end = start + 1
  while (end < records.length && records[end]!.id !== 0x000a) end++
  const bytes = drawingBytes(records, start + 1, end, DRAWING)
  const budget = { remaining: MAX_ART_RECORDS },
    images: SheetImage[] = []
  const visit = (data: Uint8Array, depth: number): void => {
    if (depth > 24) throw new ParseError("XLS drawing nesting exceeds the resource limit")
    for (const record of children(data, budget)) {
      if (record.type !== 0xf004) {
        if (record.version === 15) visit(record.data, depth + 1)
        continue
      }
      let index = 0,
        name: string | undefined
      let anchor: SheetImage["anchor"] | undefined
      for (const item of children(record.data, budget)) {
        const view = new DataView(item.data.buffer, item.data.byteOffset, item.data.byteLength)
        if (item.type === 0xf00b) {
          const size = item.instance * 6
          if (size > item.data.length) continue
          let complex = size
          for (let i = 0; i < item.instance; i++) {
            const id = view.getUint16(i * 6, true),
              value = view.getUint32(i * 6 + 2, true)
            if (id === 0x4104) index = value
            if (id & 0x8000) {
              if (value > item.data.length - complex) break
              if ((id & 0x3fff) === 0x0380 && value % 2 === 0)
                name = new TextDecoder("utf-16le")
                  .decode(item.data.subarray(complex, complex + value))
                  .replace(/\u0000+$/, "")
              complex += value
            }
          }
        } else if (item.type === 0xf010 && item.data.length === 18) {
          const from = {
            col: view.getUint16(2, true),
            colOff: view.getInt16(4, true),
            row: view.getUint16(6, true),
            rowOff: view.getInt16(8, true),
          }
          const to = {
            col: view.getUint16(10, true),
            colOff: view.getInt16(12, true),
            row: view.getUint16(14, true),
            rowOff: view.getInt16(16, true),
          }
          // fMove/fSize 表示移动或调整单元格时保持对象不动/不变形，不能一律拉伸。
          const flags = view.getUint16(0, true)
          if (from.col <= 255 && to.col <= 255 && from.row <= to.row && from.col <= to.col)
            anchor = {
              kind: "twoCell",
              editAs: flags & 1 ? "absolute" : flags & 2 ? "oneCell" : "twoCell",
              offsetUnit: "biff8",
              from,
              to,
            }
        }
      }
      const source = pool[index - 1]
      if (source && anchor)
        images.push({ ...source, anchor, ...(name ? { title: name, altText: name } : {}) })
    }
  }
  visit(bytes, 0)
  return images
}
