import { describe, expect, it } from "vitest"
import { CfbReader } from "../src/xls/cfb"
import { writeCfb } from "../src/xlsx/crypto/cfb"

const source = writeCfb([
  { name: "Small", data: Uint8Array.from({ length: 1_300 }, (_, index) => index % 251) },
  { name: "Large", data: new Uint8Array(9_000) },
])

function changed(change: (view: DataView) => void): Uint8Array {
  const data = source.slice()
  change(new DataView(data.buffer))
  return data
}

describe("CFB sector-chain validation", () => {
  it("rejects incompatible header geometry and an incomplete FAT declaration", () => {
    expect(() => new CfbReader(changed((v) => v.setUint16(0x1c, 0, true)))).toThrow(/byte order/)
    expect(() => new CfbReader(changed((v) => v.setUint16(0x1e, 8, true)))).toThrow(/sector size/)
    expect(() => new CfbReader(changed((v) => v.setUint16(0x20, 5, true)))).toThrow(
      /mini sector size/,
    )
    expect(() => new CfbReader(changed((v) => v.setUint32(0x2c, 2, true)))).toThrow(/DIFAT/)
    expect(() => new CfbReader(changed((v) => v.setUint32(0x4c, 0x1000, true)))).toThrow(
      /outside file/,
    )
  })

  it("rejects reserved, out-of-range and cyclic directory chains", () => {
    for (const sector of [0xfffffffd, 0xfffffffc]) {
      expect(() => new CfbReader(changed((v) => v.setUint32(0x30, sector, true)))).toThrow(
        /FAT\/DIFAT/,
      )
    }
    const corruptLink = (next: number) =>
      changed((v) => {
        const fatSector = v.getUint32(0x4c, true)
        const directorySector = v.getUint32(0x30, true)
        v.setUint32((fatSector + 1) * 512 + directorySector * 4, next, true)
      })
    expect(() => new CfbReader(corruptLink(0x1000))).toThrow(/outside the FAT/)
    expect(() => new CfbReader(corruptLink(0xfffffffd))).toThrow(/FAT\/DIFAT/)
    const firstDir = new DataView(source.buffer).getUint32(0x30, true)
    expect(() => new CfbReader(corruptLink(firstDir))).toThrow(/cyclic sector/)
  })

  it("rejects invalid mini-stream links without affecting ordinary streams", () => {
    const valid = new CfbReader(source)
    expect(valid.getStream("Small")?.length).toBe(1_300)
    expect(valid.getStream("Large")?.length).toBe(9_000)

    const withMiniStart = (sector: number) =>
      changed((v) => {
        const directory = v.getUint32(0x30, true)
        v.setUint32((directory + 1) * 512 + 128 + 116, sector, true)
      })
    expect(() => new CfbReader(withMiniStart(0x1000)).getStream("Small")).toThrow(
      /outside the mini FAT/,
    )
    expect(() => new CfbReader(withMiniStart(100)).getStream("Small")).toThrow(
      /outside the root mini stream/,
    )

    const cyclic = changed((v) => {
      const miniFatSector = v.getUint32(0x3c, true)
      v.setUint32((miniFatSector + 1) * 512, 0, true)
    })
    expect(() => new CfbReader(cyclic).getStream("Small")).toThrow(/cyclic mini sector/)
  })
})
