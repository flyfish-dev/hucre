import { describe, expect, it } from "vitest"
import { parseExcelRichValueImages } from "../src/xlsx/rich-value-images-reader"

const metadata = `<metadata><metadataTypes><metadataType name="XLRICHVALUE"/></metadataTypes>
  <futureMetadata name="XLRICHVALUE"><bk><ext><rvb i="0"/></ext></bk></futureMetadata>
  <valueMetadata><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`
const structures = `<rvStructures><s><k n="_rvRel:LocalImageIdentifier"/></s></rvStructures>`
const values = `<rvData><rv s="0"><v>0</v></rv></rvData>`
const relationships = `<richValueRels><rel r:id="rImage"/></richValueRels>`

describe("Excel 365 rich-value image metadata validation", () => {
  it("resolves a nested rich-value marker and namespaced relationship ID", () => {
    const result = parseExcelRichValueImages(metadata, structures, values, relationships)
    expect(result.refs).toEqual([
      { id: "excel365-rich-value-0", richValueIndex: 0, embedRId: "rImage" },
    ])
    expect([...result.imageIdByVm]).toEqual([[1, "excel365-rich-value-0"]])
  })

  it("ignores absent or unrelated parts without inventing image IDs", () => {
    for (const parts of [
      ["<other/>", structures, values, relationships],
      [metadata, "<other/>", values, relationships],
      [metadata, structures, "<other/>", relationships],
      [metadata, structures, values, "<other/>"],
      [
        metadata.replace(/<valueMetadata>.*<\/valueMetadata>/s, ""),
        structures,
        values,
        relationships,
      ],
    ]) {
      const result = parseExcelRichValueImages(...(parts as [string, string, string, string]))
      expect(result.imageIdByVm.size === 0 || result.refs.length === 0).toBe(true)
    }
  })

  it("keeps rich-value indexes stable while rejecting invalid structure and relationship indexes", () => {
    const assortedValues = `<rvData><other/><rv s="bad"><v>0</v></rv><rv s="0"><v>-1</v></rv>
      <rv s="0"><v>1.5</v></rv><rv s="0"><v>99</v></rv><rv s="0"><v>0</v></rv></rvData>`
    const result = parseExcelRichValueImages(metadata, structures, assortedValues, relationships)
    expect(result.refs).toEqual([
      { id: "excel365-rich-value-4", richValueIndex: 4, embedRId: "rImage" },
    ])
    expect(result.imageIdByVm.size).toBe(0)
  })

  it("skips malformed vm metadata and non-image rich-value structures", () => {
    const noisyMetadata = `<metadata><metadataTypes><other/><metadataType name="OTHER"/>
      <metadataType name="XLRICHVALUE"/></metadataTypes>
      <futureMetadata name="OTHER"><bk><rvb i="0"/></bk></futureMetadata>
      <futureMetadata name="XLRICHVALUE"><other/><bk><rvb i="bad"/></bk><bk><ext><rvb i="0"/></ext></bk></futureMetadata>
      <valueMetadata><other/><bk><rc t="0" v="1"/><rc t="1" v="1"/><rc t="2" v="bad"/>
        <rc t="2" v="0"/><rc t="2" v="1"/></bk></valueMetadata></metadata>`
    const result = parseExcelRichValueImages(
      noisyMetadata,
      `<rvStructures><other/><s><k n="Other"/></s><s><other/><k n="_rvRel:LocalImageIdentifier"/></s></rvStructures>`,
      `<rvData><other/><rv s="0"><v>0</v></rv><rv s="1"><v>0</v></rv></rvData>`,
      `<richValueRels><other/><rel/><rel id="rImage"/></richValueRels>`,
    )
    expect(result.refs).toEqual([
      { id: "excel365-rich-value-1", richValueIndex: 1, embedRId: "rImage" },
    ])
    expect(result.imageIdByVm.size).toBe(0)
  })
})
