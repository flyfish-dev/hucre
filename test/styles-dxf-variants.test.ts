import { describe, expect, it } from "vitest"
import { parseStyles } from "../src/xlsx/styles"

describe("XLSX sparse differential styles", () => {
  it("keeps inline and built-in number formats distinct and resolves a late indexed palette", () => {
    const styles = parseStyles(`<styleSheet>
      <dxfs>
        <dxf><font><b/></font><numFmt numFmtId="200" formatCode="0.000"/>
          <fill><patternFill patternType="solid"><fgColor indexed="0"/></patternFill></fill>
          <border><bottom style="thin"><color indexed="0"/></bottom></border>
          <alignment horizontal="center"/></dxf>
        <dxf><numFmt numFmtId="14"/></dxf>
        <dxf><numFmt numFmtId="not-a-number"/></dxf>
        <other/>
      </dxfs>
      <colors>text<other/><indexedColors>text<other/><rgbColor rgb="FF112233"/></indexedColors></colors>
    </styleSheet>`)
    expect(styles.dxfs).toHaveLength(3)
    expect(styles.dxfs[0]).toMatchObject({
      font: { bold: true },
      numFmt: "0.000",
      fill: { type: "pattern", pattern: "solid", fgColor: { indexed: 0, rgb: "112233" } },
      border: { bottom: { style: "thin", color: { indexed: 0, rgb: "112233" } } },
      alignment: { horizontal: "center" },
    })
    expect(styles.dxfs[1].numFmt).toBe("m/d/yyyy")
    expect(styles.dxfs[2].numFmt).toBeUndefined()
  })

  it("does not invent a palette or format from empty differential records", () => {
    const styles = parseStyles(`<styleSheet><dxfs><dxf><numFmt numFmtId="999"/></dxf>
      <dxf/><other/></dxfs><colors><indexedColors/></colors></styleSheet>`)
    expect(styles.dxfs).toEqual([{}, {}])
  })
})
