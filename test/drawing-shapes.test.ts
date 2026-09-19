import { describe, expect, it } from "vitest"
import { readDrawingShapePrimitives } from "../src/xlsx/drawing-shapes"
import { readDrawingLayout } from "../src/xlsx/drawing-layout"
import { parseXml } from "../src/xml/parser"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsx } from "../src/xlsx/writer"
import { cloneSheet, deleteColumns, deleteRows, insertColumns, insertRows } from "../src/sheet-ops"
import { deserializeWorkbook, serializeWorkbook } from "../src/worker"

const rectangle = `<sp><spPr><xfrm><off x="100" y="200"/><ext cx="100" cy="100"/></xfrm>
  <prstGeom prst="rect"/><noFill/><ln w="9525"><solidFill><srgbClr val="112233"/></solidFill></ln>
</spPr></sp>`
const ellipse = `<sp><spPr><xfrm><off x="125" y="225"/><ext cx="50" cy="50"/></xfrm>
  <prstGeom prst="ellipse"/><solidFill><schemeClr val="bg1"/></solidFill>
</spPr></sp>`
const group = `<grpSp><grpSpPr><xfrm><off x="5" y="10"/><ext cx="200" cy="200"/>
  <chOff x="100" y="200"/><chExt cx="100" cy="100"/></xfrm></grpSpPr>
  ${rectangle}${ellipse}</grpSp>`

describe("DrawingML basic vector shapes", () => {
  it("keeps a direct shape inside its cell anchor", () => {
    const anchor = parseXml(`<twoCellAnchor>${rectangle}</twoCellAnchor>`)
    expect(readDrawingShapePrimitives(anchor)).toEqual([
      {
        geometry: "rect",
        x: 0,
        y: 0,
        width: 1,
        height: 1,
        fill: null,
        stroke: { rgb: "112233" },
        strokeWidth: 1,
      },
    ])
  })

  it("maps group child coordinates through chOff/chExt, without using sheet offsets", () => {
    const anchor = parseXml(`<twoCellAnchor editAs="oneCell">${group}</twoCellAnchor>`)
    expect(readDrawingLayout(anchor)).toEqual({
      kind: "twoCell",
      editAs: "oneCell",
      extent: { cx: 200, cy: 200 },
    })
    expect(readDrawingShapePrimitives(anchor)).toEqual([
      {
        geometry: "rect",
        x: 0,
        y: 0,
        width: 1,
        height: 1,
        fill: null,
        stroke: { rgb: "112233" },
        strokeWidth: 1,
      },
      { geometry: "ellipse", x: 0.25, y: 0.25, width: 0.5, height: 0.5, fill: { scheme: "bg1" } },
    ])
  })

  it("ignores unknown geometry and invalid group transforms", () => {
    expect(readDrawingShapePrimitives(parseXml("<twoCellAnchor><pic/></twoCellAnchor>"))).toEqual(
      [],
    )
    expect(
      readDrawingShapePrimitives(
        parseXml('<twoCellAnchor><sp><spPr><prstGeom prst="star5"/></spPr></sp></twoCellAnchor>'),
      ),
    ).toEqual([])
    expect(
      readDrawingShapePrimitives(
        parseXml(
          '<twoCellAnchor><grpSp><grpSpPr><xfrm><chExt cx="0" cy="1"/></xfrm></grpSpPr></grpSp></twoCellAnchor>',
        ),
      ),
    ).toEqual([])
  })

  it("bounds nested groups and rejects non-integer or non-positive child transforms", () => {
    const wrapper = (inside: string, off = 'x="0" y="0"', ext = 'cx="10" cy="10"') =>
      `<grpSp><grpSpPr><xfrm><off ${off}/><ext ${ext}/><chOff x="0" y="0"/><chExt cx="10" cy="10"/></xfrm></grpSpPr>${inside}</grpSp>`
    for (const off of ['x="NaN" y="0"', 'x="9007199254740992" y="0"']) {
      expect(
        readDrawingShapePrimitives(
          parseXml(`<twoCellAnchor>${wrapper(wrapper(rectangle, off))}</twoCellAnchor>`),
        ),
      ).toEqual([])
    }
    expect(
      readDrawingShapePrimitives(
        parseXml(
          `<twoCellAnchor>${wrapper(wrapper(rectangle, 'x="0" y="0"', 'cx="0" cy="10"'))}</twoCellAnchor>`,
        ),
      ),
    ).toEqual([])
    let nested = rectangle
    for (let i = 0; i < 33; i++) nested = wrapper(nested)
    expect(
      readDrawingShapePrimitives(parseXml(`<twoCellAnchor>${nested}</twoCellAnchor>`)),
    ).toEqual([])
  })

  it("retains safe system colors while ignoring unsupported color and line attributes", () => {
    const shape = parseXml(`<twoCellAnchor><sp><spPr><prstGeom prst="ellipse"/>
      <solidFill><sysClr lastClr="123ABC"/></solidFill><ln w="bad"><solidFill><schemeClr val="Bad-Color"/></solidFill></ln>
      </spPr></sp></twoCellAnchor>`)
    expect(readDrawingShapePrimitives(shape)).toEqual([
      {
        geometry: "ellipse",
        x: 0,
        y: 0,
        width: 1,
        height: 1,
        fill: { rgb: "123ABC" },
      },
    ])
  })

  it("reads an anonymous shape alongside a picture and retains it across clone/worker", async () => {
    const original = await writeXlsx({
      sheets: [
        {
          name: "S",
          rows: [["a"]],
          images: [
            {
              type: "png",
              data: new Uint8Array([137, 80, 78, 71]),
              anchor: { from: { row: 0, col: 0 }, to: { row: 1, col: 1 } },
            },
          ],
        },
      ],
    })
    const zip = new ZipReader(original)
    const rebuilt = new ZipWriter()
    for (const path of zip.entries()) {
      const bytes = await zip.extract(path)
      if (path !== "xl/drawings/drawing1.xml") {
        rebuilt.add(path, bytes)
        continue
      }
      const drawing = new TextDecoder().decode(bytes)
      const shape = `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp><xdr:spPr><a:prstGeom prst="ellipse"/><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></xdr:spPr></xdr:sp>
        <xdr:clientData/></xdr:twoCellAnchor>`
      const oneCell = `<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:from>
        <xdr:ext cx="952500" cy="190500"/><xdr:sp><xdr:spPr><a:prstGeom prst="rect"/><a:noFill/></xdr:spPr></xdr:sp>
        <xdr:clientData/></xdr:oneCellAnchor>`
      const absolute = `<xdr:absoluteAnchor><xdr:pos x="95250" y="190500"/><xdr:ext cx="190500" cy="95250"/>
        <xdr:sp><xdr:spPr><a:prstGeom prst="ellipse"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></xdr:spPr></xdr:sp>
        <xdr:clientData/></xdr:absoluteAnchor>`
      rebuilt.add(
        path,
        new TextEncoder().encode(
          drawing.replace("</xdr:wsDr>", `${shape}${oneCell}${absolute}</xdr:wsDr>`),
        ),
      )
    }
    const workbook = await readXlsx(await rebuilt.build())
    expect(workbook.sheets[0].images).toHaveLength(1)
    expect(workbook.sheets[0].shapes).toEqual([
      {
        anchor: { kind: "twoCell", from: { row: 2, col: 1 }, to: { row: 3, col: 2 } },
        primitives: [
          { geometry: "ellipse", x: 0, y: 0, width: 1, height: 1, fill: { rgb: "ABCDEF" } },
        ],
      },
      {
        anchor: { kind: "oneCell", from: { row: 4, col: 3 }, extent: { cx: 952500, cy: 190500 } },
        primitives: [{ geometry: "rect", x: 0, y: 0, width: 1, height: 1, fill: null }],
      },
      {
        anchor: {
          kind: "absolute",
          from: { row: 0, col: 0 },
          position: { x: 95250, y: 190500 },
          extent: { cx: 190500, cy: 95250 },
        },
        primitives: [
          { geometry: "ellipse", x: 0, y: 0, width: 1, height: 1, fill: { scheme: "accent1" } },
        ],
      },
    ])
    expect(cloneSheet(workbook.sheets[0], "Copy").shapes).toEqual(workbook.sheets[0].shapes)
    expect(deserializeWorkbook(serializeWorkbook(workbook)).sheets[0].shapes).toEqual(
      workbook.sheets[0].shapes,
    )
  })

  it("moves and removes vector anchors with the same row and column edits as pictures", () => {
    const sheet = {
      name: "Vectors",
      rows: Array.from({ length: 5 }, () => Array(5).fill(null)),
      shapes: [
        {
          anchor: { from: { row: 2, col: 2 }, to: { row: 3, col: 3 } },
          primitives: [{ geometry: "rect" as const, x: 0, y: 0, width: 1, height: 1 }],
        },
      ],
    }
    insertRows(sheet, 1, 2)
    insertColumns(sheet, 1, 2)
    expect(sheet.shapes[0].anchor).toEqual({ from: { row: 4, col: 4 }, to: { row: 5, col: 5 } })
    deleteRows(sheet, 0, 1)
    deleteColumns(sheet, 0, 1)
    expect(sheet.shapes[0].anchor).toEqual({ from: { row: 3, col: 3 }, to: { row: 4, col: 4 } })
    deleteRows(sheet, 3, 1)
    expect(sheet.shapes).toEqual([])
  })
})
