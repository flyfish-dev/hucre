import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { ZipWriter } from "../src/zip/writer"

const encoder = new TextEncoder()
const bytes = (seed: number) =>
  new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, seed])

function xml(value: string): Uint8Array {
  return encoder.encode(value)
}

async function buildExcel365ImageWorkbook(): Promise<Uint8Array> {
  const zip = new ZipWriter()
  zip.add(
    "[Content_Types].xml",
    xml(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`),
  )
  zip.add(
    "_rels/.rels",
    xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
  )
  zip.add(
    "xl/workbook.xml",
    xml(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Images" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
  )
  zip.add(
    "xl/_rels/workbook.xml.rels",
    xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
</Relationships>`),
  )
  zip.add(
    "xl/metadata.xml",
    xml(`<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>
  <valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>
</metadata>`),
  )
  zip.add(
    "xl/richData/rdrichvaluestructure.xml",
    xml(`<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>
</rvStructures>`),
  )
  zip.add(
    "xl/richData/rdrichvalue.xml",
    xml(
      `<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>`,
    ),
  )
  zip.add(
    "xl/richData/richValueRel.xml",
    xml(
      `<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rel r:id="rIdCell"/></richValueRels>`,
    ),
  )
  zip.add(
    "xl/richData/_rels/richValueRel.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCell" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/in-cell.png"/></Relationships>`,
    ),
  )
  zip.add(
    "xl/worksheets/sheet1.xml",
    xml(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Image</t></is></c></row><row r="2"><c r="B2" t="e" vm="1"><v>#VALUE!</v></c></row></sheetData><drawing r:id="rIdDrawing"/>
</worksheet>`),
  )
  zip.add(
    "xl/worksheets/_rels/sheet1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    ),
  )
  zip.add(
    "xl/drawings/drawing1.xml",
    xml(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>4</xdr:col><xdr:colOff>636399</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>77884</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>7</xdr:col><xdr:colOff>36697</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>125312</xdr:rowOff></xdr:to>
    <mc:AlternateContent><mc:Choice Requires="future"><xdr:graphicFrame/></mc:Choice><mc:Fallback>
      <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Fallback preview" descr="Preview"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdPreview" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:ext cx="2257798" cy="2562028"/></a:xfrm></xdr:spPr></xdr:pic>
    </mc:Fallback></mc:AlternateContent><xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
  )
  zip.add(
    "xl/drawings/_rels/drawing1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/floating.png"/></Relationships>`,
    ),
  )
  zip.add("xl/media/in-cell.png", bytes(1))
  zip.add("xl/media/floating.png", bytes(2))
  return zip.build()
}

describe("Excel 365 in-cell and AlternateContent pictures", () => {
  it("resolves vm metadata and preserves exact drawing offsets", async () => {
    const workbook = await readXlsx(await buildExcel365ImageWorkbook())
    expect(workbook.cellImages).toHaveLength(1)
    expect(workbook.cellImages?.[0]).toMatchObject({
      id: "excel365-rich-value-0",
      type: "png",
    })
    expect(workbook.cellImages?.[0]?.data.at(-1)).toBe(1)

    const sheet = workbook.sheets[0]!
    expect(sheet.cells?.get("1,1")?.imageId).toBe("excel365-rich-value-0")
    expect(sheet.images).toHaveLength(1)
    expect(sheet.images?.[0]?.data.at(-1)).toBe(2)
    expect(sheet.images?.[0]?.anchor).toEqual({
      kind: "twoCell",
      from: { row: 8, col: 4, rowOff: 77884, colOff: 636399 },
      to: { row: 19, col: 7, rowOff: 125312, colOff: 36697 },
    })
    expect(sheet.images?.[0]).toMatchObject({ width: 237, height: 269, altText: "Preview" })
  })
})
