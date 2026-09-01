import { describe, it, expect } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { createStylesCollector } from "../src/xlsx/styles-writer"
import { createSharedStrings, writeWorksheetXml } from "../src/xlsx/worksheet-writer"
import { parseXml } from "../src/xml/parser"
import type { WriteSheet } from "../src/_types"
import { parseWorksheet } from "../src/xlsx/worksheet"
import type { WorksheetContext } from "../src/xlsx/worksheet"

// ── Helpers ──────────────────────────────────────────────────────────

function writeXml(sheet: WriteSheet): string {
  const styles = createStylesCollector()
  const ss = createSharedStrings()
  const result = writeWorksheetXml(sheet, styles, ss)
  return result.xml
}

function findChild(el: { children: Array<unknown> }, localName: string): any {
  return el.children.find((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

function findChildren(el: { children: Array<unknown> }, localName: string): any[] {
  return el.children.filter((c: any) => typeof c !== "string" && (c.local || c.tag) === localName)
}

const worksheetContext: WorksheetContext = {
  sharedStrings: [],
  styles: null,
  readStyles: false,
  dateSystem: "1900",
}

// ── Row Breaks Writing ──────────────────────────────────────────────

describe("page breaks — row breaks writing", () => {
  it("writes row breaks with correct XML structure", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"], ["B"], ["C"]],
      rowBreaks: [9, 24], // 0-based: break after row 10 and 25
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    const rb = findChild(doc, "rowBreaks")
    expect(rb).toBeDefined()
    expect(rb.attrs["count"]).toBe("2")
    expect(rb.attrs["manualBreakCount"]).toBe("2")

    const brks = findChildren(rb, "brk")
    expect(brks).toHaveLength(2)

    // id = 0-based + 1 = 1-based
    expect(brks[0].attrs["id"]).toBe("10")
    expect(brks[0].attrs["max"]).toBe("16383")
    expect(brks[0].attrs["man"]).toBe("1")

    expect(brks[1].attrs["id"]).toBe("25")
    expect(brks[1].attrs["max"]).toBe("16383")
    expect(brks[1].attrs["man"]).toBe("1")
  })

  it("sorts row breaks in output", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      rowBreaks: [24, 9], // unsorted
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    const rb = findChild(doc, "rowBreaks")
    const brks = findChildren(rb, "brk")

    // Should be sorted: 10, 25 (1-based)
    expect(brks[0].attrs["id"]).toBe("10")
    expect(brks[1].attrs["id"]).toBe("25")
  })
})

// ── Column Breaks Writing ───────────────────────────────────────────

describe("page breaks — column breaks writing", () => {
  it("writes column breaks with correct XML structure", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      colBreaks: [4], // 0-based: break after column E
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    const cb = findChild(doc, "colBreaks")
    expect(cb).toBeDefined()
    expect(cb.attrs["count"]).toBe("1")
    expect(cb.attrs["manualBreakCount"]).toBe("1")

    const brks = findChildren(cb, "brk")
    expect(brks).toHaveLength(1)

    expect(brks[0].attrs["id"]).toBe("5") // 0-based + 1
    expect(brks[0].attrs["max"]).toBe("1048575")
    expect(brks[0].attrs["man"]).toBe("1")
  })
})

// ── Both Row and Column Breaks ──────────────────────────────────────

describe("page breaks — both row and column breaks", () => {
  it("writes both row and column breaks", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      rowBreaks: [9],
      colBreaks: [4],
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    expect(findChild(doc, "rowBreaks")).toBeDefined()
    expect(findChild(doc, "colBreaks")).toBeDefined()
  })
})

// ── No Breaks ───────────────────────────────────────────────────────

describe("page breaks — no breaks", () => {
  it("omits rowBreaks/colBreaks elements when no breaks defined", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    expect(findChild(doc, "rowBreaks")).toBeUndefined()
    expect(findChild(doc, "colBreaks")).toBeUndefined()
  })

  it("omits elements for empty arrays", () => {
    const sheet: WriteSheet = {
      name: "Test",
      rows: [["A"]],
      rowBreaks: [],
      colBreaks: [],
    }

    const xml = writeXml(sheet)
    const doc = parseXml(xml)

    expect(findChild(doc, "rowBreaks")).toBeUndefined()
    expect(findChild(doc, "colBreaks")).toBeUndefined()
  })
})

// ── Round-trip (write → read) ───────────────────────────────────────

describe("page breaks — round-trip", () => {
  it("row breaks survive write → read cycle", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["Row 1"], ["Row 2"], ["Row 3"]],
          rowBreaks: [0, 1],
        },
      ],
    })

    const wb = await readXlsx(xlsx)
    expect(wb.sheets[0].rowBreaks).toBeDefined()
    expect(wb.sheets[0].rowBreaks).toEqual([0, 1])
  })

  it("column breaks survive write → read cycle", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A", "B", "C"]],
          colBreaks: [1],
        },
      ],
    })

    const wb = await readXlsx(xlsx)
    expect(wb.sheets[0].colBreaks).toBeDefined()
    expect(wb.sheets[0].colBreaks).toEqual([1])
  })

  it("both row and column breaks survive round-trip", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"]],
          rowBreaks: [9, 24],
          colBreaks: [4, 7],
        },
      ],
    })

    const wb = await readXlsx(xlsx)
    expect(wb.sheets[0].rowBreaks).toEqual([9, 24])
    expect(wb.sheets[0].colBreaks).toEqual([4, 7])
  })

  it("multiple breaks are sorted after reading", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"]],
          rowBreaks: [24, 9, 15],
        },
      ],
    })

    const wb = await readXlsx(xlsx)
    // Should come back sorted
    expect(wb.sheets[0].rowBreaks).toEqual([9, 15, 24])
  })

  it("preserves first and final-grid break boundaries", async () => {
    const xlsx = await writeXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: [["A"]],
          rowBreaks: [0, 1_048_574],
          colBreaks: [0, 16_382],
        },
      ],
    })

    const sheet = (await readXlsx(xlsx)).sheets[0]!
    expect(sheet.rowBreaks).toEqual([0, 1_048_574])
    expect(sheet.colBreaks).toEqual([0, 16_382])
  })

  it("deduplicates breaks and omits invalid or out-of-grid positions", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [["A"]],
      rowBreaks: [-1, 0, 0, 1.5, Number.NaN, 1_048_574, 1_048_575],
      colBreaks: [-1, 0, 0, 1.5, Number.POSITIVE_INFINITY, 16_382, 16_383],
    }
    const xml = writeXml(sheet)
    const doc = parseXml(xml)
    const rowIds = findChildren(findChild(doc, "rowBreaks"), "brk").map(
      (entry) => entry.attrs["id"],
    )
    const colIds = findChildren(findChild(doc, "colBreaks"), "brk").map(
      (entry) => entry.attrs["id"],
    )

    expect(rowIds).toEqual(["1", "1048575"])
    expect(colIds).toEqual(["1", "16383"])

    const roundTripped = (await readXlsx(await writeXlsx({ sheets: [sheet] }))).sheets[0]!
    expect(roundTripped.rowBreaks).toEqual([0, 1_048_574])
    expect(roundTripped.colBreaks).toEqual([0, 16_382])
  })

  it("keeps break collections absent when every supplied position is invalid", async () => {
    const sheet: WriteSheet = {
      name: "Sheet1",
      rows: [],
      rowBreaks: [-1, 1.5, 1_048_575],
      colBreaks: [-1, Number.NaN, 16_383],
    }
    const doc = parseXml(writeXml(sheet))

    expect(findChild(doc, "rowBreaks")).toBeUndefined()
    expect(findChild(doc, "colBreaks")).toBeUndefined()

    const roundTripped = (await readXlsx(await writeXlsx({ sheets: [sheet] }))).sheets[0]!
    expect(roundTripped.rowBreaks).toBeUndefined()
    expect(roundTripped.colBreaks).toBeUndefined()
  })

  it("normalizes malformed SpreadsheetML break ids while reading", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData/>
        <rowBreaks count="7" manualBreakCount="7">
          <brk id="0"/><brk id="1"/><brk id="1"/><brk id="1.5"/>
          <brk id="1048575"/><brk id="1048576"/><brk id="invalid"/>
        </rowBreaks>
        <colBreaks count="7" manualBreakCount="7">
          <brk id="0"/><brk id="1"/><brk id="1"/><brk id="1.5"/>
          <brk id="16383"/><brk id="16384"/><brk id="invalid"/>
        </colBreaks>
      </worksheet>`
    const sheet = parseWorksheet(xml, "Sheet1", worksheetContext)

    expect(sheet.rowBreaks).toEqual([0, 1_048_574])
    expect(sheet.colBreaks).toEqual([0, 16_382])
  })
})
