import { describe, expect, it } from "vitest"
import { parseXml } from "../src/xml/parser"
import { readDrawingLayout } from "../src/xlsx/drawing-layout"
import { parseStyles } from "../src/xlsx/styles"

function layout(xml: string) {
  return readDrawingLayout(parseXml(xml))
}

function picture(extent: string, position = "") {
  return `<pic><spPr><xfrm>${position}${extent}</xfrm></spPr></pic>`
}

const extent = '<ext cx="1252822" cy="962025"/>'

describe("DrawingML metadata ownership", () => {
  it("ignores unrelated containers", () => {
    expect(layout(`<graphicFrame>${picture(extent)}</graphicFrame>`)).toEqual({})
  })

  it("does not invent an extent or behavior for incomplete anchors", () => {
    expect(layout('<twoCellAnchor editAs="unknown">text<from/></twoCellAnchor>')).toEqual({
      kind: "twoCell",
    })
    expect(layout(`<twoCellAnchor>${picture("")}</twoCellAnchor>`)).toEqual({
      kind: "twoCell",
    })
  })

  for (const editAs of ["twoCell", "oneCell", "absolute"]) {
    it(`retains the explicit ${editAs} behavior`, () => {
      const xml = `<twoCellAnchor editAs="${editAs}">${picture(extent)}</twoCellAnchor>`
      expect(layout(xml)).toEqual({
        kind: "twoCell",
        editAs,
        extent: { cx: 1252822, cy: 962025 },
      })
    })
  }

  it("takes one-cell extents from the container rather than a stale picture transform", () => {
    const xml = `<oneCellAnchor><ext cx="0" cy="1"/>${picture(extent)}</oneCellAnchor>`
    expect(layout(xml)).toEqual({ kind: "oneCell", extent: { cx: 0, cy: 1 } })
  })

  it("keeps signed absolute coordinates without floating-point pixel conversion", () => {
    expect(layout('<absoluteAnchor><pos x="-3175" y="19051"/></absoluteAnchor>')).toEqual({
      kind: "absolute",
      position: { x: -3175, y: 19051 },
    })
    const xml = picture(extent, '<off x="10" y="20"/>')
    expect(layout(`<twoCellAnchor editAs="absolute">${xml}</twoCellAnchor>`).position).toEqual({
      x: 10,
      y: 20,
    })
    expect(layout(`<oneCellAnchor>${xml}</oneCellAnchor>`).position).toBeUndefined()
  })

  for (const attribute of ["cx", "cy"]) {
    for (const value of ["", "NaN", "Infinity", "1.5", "-1", "9007199254740992"]) {
      it(`rejects invalid ${attribute}=${value}`, () => {
        const ext = extent.replace(new RegExp(`${attribute}="[^"]*"`), `${attribute}="${value}"`)
        expect(layout(`<twoCellAnchor>${picture(ext)}</twoCellAnchor>`).extent).toBeUndefined()
      })
    }
  }

  for (const attrs of ['cx="1"', 'cy="1"', ""]) {
    it(`rejects incomplete extent ${attrs}`, () => {
      const xml = `<oneCellAnchor><ext ${attrs}/></oneCellAnchor>`
      expect(layout(xml).extent).toBeUndefined()
    })
  }

  it("accepts valid signed and whitespace-padded integer coordinates", () => {
    const xml = '<oneCellAnchor><ext cx=" +1 " cy=" 2 "/></oneCellAnchor>'
    expect(layout(xml).extent).toEqual({ cx: 1, cy: 2 })
  })

  it("finds fallback pictures through compatibility wrappers, not grouped transforms", () => {
    const wrapped = `<AlternateContent>\n<Choice><graphicFrame/></Choice>\n<Fallback>${picture(extent)}</Fallback></AlternateContent>`
    expect(layout(`<twoCellAnchor>${wrapped}</twoCellAnchor>`).extent).toEqual({
      cx: 1252822,
      cy: 962025,
    })
    const group = `<grpSp>${picture(extent)}</grpSp>`
    expect(layout(`<twoCellAnchor>${group}</twoCellAnchor>`).extent).toBeUndefined()
    expect(layout(`<twoCellAnchor>text<AlternateContent/></twoCellAnchor>`).extent).toBeUndefined()
  })

  it("supports local-name fallback on caller-created XML nodes", () => {
    const node = parseXml(`<twoCellAnchor>${picture(extent)}</twoCellAnchor>`)
    node.local = ""
    expect(readDrawingLayout(node).extent).toEqual({ cx: 1252822, cy: 962025 })
  })
})

function styles(xf = "1", font = "1") {
  return parseStyles(`<styleSheet>
    <fonts><font><name val="Fallback"/></font><font><name val="Normal"/></font></fonts>
    <cellStyleXfs><xf fontId="0"/><xf fontId="${font}"/></cellStyleXfs>
    <cellStyles><cellStyle builtinId="0" xfId="${xf}"/></cellStyles>
  </styleSheet>`)
}

describe("Normal style indirection", () => {
  it("uses builtin Normal rather than font table order", () => {
    expect(styles().normalFont?.name).toBe("Normal")
  })

  for (const value of ["-1", "1.5", "999", "NaN"]) {
    it(`falls back for an invalid style or font reference ${value}`, () => {
      expect(styles(value).normalFont?.name).toBe("Fallback")
      expect(styles("1", value).normalFont?.name).toBe("Fallback")
    })
  }

  it("tolerates missing styles and fonts", () => {
    expect(parseStyles("<styleSheet/>").normalFont).toBeUndefined()
    const xml = '<styleSheet><fonts><font><name val="Only"/></font></fonts></styleSheet>'
    expect(parseStyles(xml).normalFont?.name).toBe("Only")
  })
})
