import { describe, expect, it } from "vitest"
import { readXlsPicturePool, readXlsSheetImages } from "../src/xls/drawings"
import { writeDrawing } from "../src/xlsx/drawing-writer"
import type { BiffRecord } from "../src/xls/biff"

const concat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((size, part) => size + part.length, 0))
  let cursor = 0
  for (const part of parts) {
    out.set(part, cursor)
    cursor += part.length
  }
  return out
}
function art(type: number, body: Uint8Array, instance = 0, version = 0): Uint8Array {
  const bytes = new Uint8Array(body.length + 8),
    view = new DataView(bytes.buffer)
  view.setUint16(0, (instance << 4) | version, true)
  view.setUint16(2, type, true)
  view.setUint32(4, body.length, true)
  bytes.set(body, 8)
  return bytes
}
const png = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jL1cAAAAASUVORK5CYII=",
  ),
  (c) => c.charCodeAt(0),
)
function bse(valid = true, two = false): Uint8Array {
  const uid = new Uint8Array(two ? 33 : 17)
  const blip = art(valid ? 0xf01e : 0xf01c, concat(uid, png), two ? 0x6e1 : 0x6e0)
  const fixed = new Uint8Array(36)
  fixed[0] = 6
  return art(0xf007, concat(fixed, blip), 6, 2)
}
const rec = (id: number, data: Uint8Array): BiffRecord => ({ id, data, offset: 0 })
function poolRecords(entries: Uint8Array[], split = false): BiffRecord[] {
  const bytes = art(0xf000, art(0xf001, concat(...entries), entries.length, 15), 0, 15)
  return split
    ? [rec(0xeb, bytes.slice(0, 19)), rec(0xeb, bytes.slice(19, 50)), rec(0x3c, bytes.slice(50))]
    : [rec(0xeb, bytes)]
}
function shape(index = 1, flags = 2, signedOffset = 512): Uint8Array {
  const properties = new Uint8Array(6),
    p = new DataView(properties.buffer)
  p.setUint16(0, 0x4104, true)
  p.setUint32(2, index, true)
  const anchor = new Uint8Array(18),
    v = new DataView(anchor.buffer)
  ;[flags, 1, signedOffset, 2, 128, 3, 256, 5, 64].forEach((n, i) => v.setUint16(i * 2, n, true))
  return art(0xf004, concat(art(0xf00b, properties, 1, 3), art(0xf010, anchor)), 0, 15)
}

describe("BIFF8 OfficeArt 图片与原生单元格锚点", () => {
  it("跨绘图记录与 CONTINUE 读取 PNG，两 UID 不吞掉图片头", () => {
    for (const two of [false, true]) {
      const records = poolRecords([bse(true, two)], true),
        pictures = readXlsPicturePool(records, records.length)
      expect(pictures).toHaveLength(1)
      expect(pictures[0]?.data).toEqual(png)
      expect(pictures[0]?.type).toBe("png")
    }
  })
  it("不支持的图片保留池位置，不能让后续引用错位", () => {
    const records = poolRecords([bse(false), bse()]),
      pictures = readXlsPicturePool(records, 1)
    expect(pictures[0]).toBeNull()
    expect(pictures[1]?.data).toEqual(png)
    const sheet = [
      rec(0x809, new Uint8Array()),
      rec(0xec, art(0xf002, concat(shape(1), shape(2)), 0, 15)),
      rec(0xa, new Uint8Array()),
    ]
    const images = readXlsSheetImages(sheet, 0, pictures)
    expect(images).toHaveLength(1)
    expect(images[0]?.anchor).toEqual({
      kind: "twoCell",
      editAs: "oneCell",
      offsetUnit: "biff8",
      from: { col: 1, colOff: 512, row: 2, rowOff: 128 },
      to: { col: 3, colOff: 256, row: 5, rowOff: 64 },
    })
  })
  it("多工作表复用同一图片字节，不复用各自的锚点对象", () => {
    const pictures = readXlsPicturePool(poolRecords([bse()]), 1)
    const sheet = [
      rec(0x809, new Uint8Array()),
      rec(0xec, art(0xf002, shape(), 0, 15)),
      rec(0xa, new Uint8Array()),
    ]
    const first = readXlsSheetImages(sheet, 0, pictures),
      second = readXlsSheetImages(sheet, 0, pictures)
    expect(first[0]?.data).toBe(second[0]?.data)
    expect(first[0]?.anchor).not.toBe(second[0]?.anchor)
  })
  it("截断、越界长度和空绘图不会猜测图片或越界访问", () => {
    const invalid = poolRecords([bse()])
    invalid[0]!.data = invalid[0]!.data.slice(0, 25)
    expect(readXlsPicturePool(invalid, 1)).toEqual([])
    expect(readXlsSheetImages([rec(0xa, new Uint8Array())], 0, [])).toEqual([])
  })
  it("XLS 图片写为 XLSX 时转换相对偏移，不把 BIFF 分数写成 EMU", () => {
    const images = readXlsSheetImages(
      [
        rec(0x809, new Uint8Array()),
        rec(0xec, art(0xf002, shape(), 0, 15)),
        rec(0xa, new Uint8Array()),
      ],
      0,
      readXlsPicturePool(poolRecords([bse()]), 1),
    )
    const output = writeDrawing(images, 1, undefined, undefined, 1, {
      columns: [{}, { width: 10 }],
      rowDefs: new Map([[2, { height: 24 }]]),
    })
    expect(output.drawingXml).toContain("<xdr:colOff>333375</xdr:colOff>")
    expect(output.drawingXml).toContain("<xdr:rowOff>152400</xdr:rowOff>")
    expect(images[0]!.anchor.offsetUnit).toBe("biff8")
    expect(images[0]!.anchor.from.colOff).toBe(512)
  })
  it("保留 fMove/fSize 与有符号相对偏移，不能在缩放时把固定图片拉伸", () => {
    const pool = readXlsPicturePool(poolRecords([bse()]), 1)
    for (const [flags, editAs] of [
      [0, "twoCell"],
      [2, "oneCell"],
      [3, "absolute"],
    ] as const) {
      const images = readXlsSheetImages(
        [
          rec(0x809, new Uint8Array()),
          rec(0xec, shape(1, flags, -512)),
          rec(0xa, new Uint8Array()),
        ],
        0,
        pool,
      )
      expect(images[0]!.anchor.editAs).toBe(editAs)
      expect(images[0]!.anchor.from.colOff).toBe(-512)
      const output = writeDrawing(images, 1, undefined, undefined, 1, {})
      expect(output.drawingXml).toContain(`editAs="${editAs}"`)
      expect(output.drawingXml).not.toContain("NaN")
      expect(images[0]!.anchor.offsetUnit).toBe("biff8")
    }
  })
  it("JPEG 的单/双 UID、PNG 签名与不支持记录都按真实头校验", () => {
    const entry = (type: number, instance: number, bytes: Uint8Array, version = 0) =>
      art(0xf007, concat(new Uint8Array(36), art(type, bytes, instance, version)), 0, 2)
    for (const [type, instance] of [
      [0xf01d, 0x46a],
      [0xf01d, 0x46b],
      [0xf02a, 0x6e2],
      [0xf02a, 0x6e3],
    ]) {
      const jpeg = new Uint8Array([255, 216, 255, 224])
      const bytes = concat(new Uint8Array(instance! & 1 ? 33 : 17), jpeg)
      expect(readXlsPicturePool(poolRecords([entry(type!, instance!, bytes)]), 1)[0]).toEqual({
        type: "jpeg",
        data: jpeg,
      })
      for (const bad of [
        new Uint8Array(),
        new Uint8Array([0, 216, 255]),
        new Uint8Array([255, 0, 255]),
        new Uint8Array([255, 216, 0]),
      ]) {
        expect(
          readXlsPicturePool(
            poolRecords([
              entry(type!, instance!, concat(new Uint8Array(instance! & 1 ? 33 : 17), bad)),
            ]),
            1,
          ),
        ).toEqual([null])
      }
    }
    expect(readXlsPicturePool(poolRecords([entry(0xf01e, 0x6e0, new Uint8Array(20))]), 1)).toEqual([
      null,
    ])
    const invalidPng = png.slice()
    invalidPng[0] = 0
    expect(
      readXlsPicturePool(
        poolRecords([entry(0xf01e, 0x6e0, concat(new Uint8Array(17), invalidPng))]),
        1,
      ),
    ).toEqual([null])
    expect(readXlsPicturePool(poolRecords([entry(0xf01e, 0x6e0, png, 2)]), 1)).toEqual([null])
    expect(readXlsPicturePool(poolRecords([art(0xf007, new Uint8Array(3))]), 1)).toEqual([null])
  })
  it("复杂图片名称、非图片属性、错误属性长度与无效端点不会串位", () => {
    const pool = readXlsPicturePool(poolRecords([bse()]), 1)
    const read = (content: Uint8Array) =>
      readXlsSheetImages(
        [
          rec(0x809, new Uint8Array()),
          rec(0xec, art(0xf004, content, 0, 15)),
          rec(0xa, new Uint8Array()),
        ],
        0,
        pool,
      )
    const basic = shape().subarray(8),
      anchor = basic.subarray(14)
    const name = new Uint8Array([78, 0, 101, 0, 117, 0, 0, 0])
    const properties = new Uint8Array(12),
      view = new DataView(properties.buffer)
    view.setUint16(0, 0x4104, true)
    view.setUint32(2, 1, true)
    view.setUint16(6, 0x8380, true)
    view.setUint32(8, name.length, true)
    expect(read(concat(art(0xf00b, concat(properties, name), 2, 3), anchor))[0]!.title).toBe("Neu")
    view.setUint32(8, 0xffff, true)
    expect(
      read(concat(art(0xf00b, concat(properties, name), 2, 3), anchor))[0]!.title,
    ).toBeUndefined()
    view.setUint32(8, 3, true)
    expect(
      read(concat(art(0xf00b, concat(properties, name), 2, 3), anchor))[0]!.title,
    ).toBeUndefined()
    expect(read(concat(art(0xf00b, new Uint8Array(2), 3, 3), anchor))).toEqual([])
    expect(read(concat(art(0xf00b, new Uint8Array(6), 1, 3), anchor))).toEqual([])
    const invalidAnchor = anchor.slice()
    new DataView(invalidAnchor.buffer).setUint16(10, 256, true)
    expect(read(concat(basic.subarray(0, 14), invalidAnchor))).toEqual([])
    expect(read(concat(basic.subarray(0, 14), art(0xf010, new Uint8Array(17))))).toEqual([])
    expect(read(concat(basic, art(0xf006, new Uint8Array())))).toHaveLength(1)
  })
  it("绘图字节、嵌套和记录数量均有硬上限", () => {
    const large = new Uint8Array(64 * 1024 * 1024)
    expect(() => readXlsPicturePool([rec(0xeb, large), rec(0x3c, new Uint8Array(1))], 2)).toThrow(
      /resource limit/,
    )
    let nested = art(0xf006, new Uint8Array())
    for (let i = 0; i < 26; i++) nested = art(0xf002, nested, 0, 15)
    expect(() => readXlsPicturePool([rec(0xeb, nested)], 1)).toThrow(/nesting/)
    const pool = readXlsPicturePool(poolRecords([bse()]), 1)
    expect(() =>
      readXlsSheetImages(
        [rec(0x809, new Uint8Array()), rec(0xec, nested), rec(0xa, new Uint8Array())],
        0,
        pool,
      ),
    ).toThrow(/nesting/)
    const count = new Uint8Array(8 * 100001)
    expect(() => readXlsPicturePool([rec(0xeb, count)], 1)).toThrow(/record count/)
    expect(() =>
      writeDrawing(
        [{ type: "png", data: png, anchor: { offsetUnit: "biff8", from: { row: 0, col: 0 } } }],
        1,
      ),
    ).toThrow(/source worksheet/)
    const noEnd = writeDrawing(
      [{ type: "png", data: png, anchor: { offsetUnit: "biff8", from: { row: 0, col: 0 } } }],
      1,
      undefined,
      undefined,
      1,
      {},
    )
    expect(noEnd.images).toHaveLength(1)
    expect(noEnd.drawingXml).not.toContain("NaN")
  })
})
