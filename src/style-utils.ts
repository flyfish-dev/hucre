import type { BorderLineStyle, Color, FillPattern } from "./_types"

export const BUILTIN_NUM_FMTS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: "$#,##0_);\\($#,##0\\)",
  6: "$#,##0_);[Red]\\($#,##0\\)",
  7: "$#,##0.00_);\\($#,##0.00\\)",
  8: "$#,##0.00_);[Red]\\($#,##0.00\\)",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
}

export const DATE_FMT_IDS: Set<number> = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58,
])

export function getBuiltinNumberFormat(id: number): string | undefined {
  return BUILTIN_NUM_FMTS[id]
}

export function rgbHex(red: number, green: number, blue: number): string {
  return [red, green, blue].map((v) => v.toString(16).padStart(2, "0").toUpperCase()).join("")
}

export function indexedColor(index: number, palette?: Map<number, Color>): Color | undefined {
  if (index === 0x7fff || index === 0x40 || index === 0x41) return undefined
  return palette?.get(index) ?? { indexed: index }
}

export function biffBorderLineStyle(code: number): BorderLineStyle | undefined {
  switch (code) {
    case 0x01:
      return "thin"
    case 0x02:
      return "medium"
    case 0x03:
      return "dashed"
    case 0x04:
      return "dotted"
    case 0x05:
      return "thick"
    case 0x06:
      return "double"
    case 0x07:
      return "hair"
    case 0x08:
      return "mediumDashed"
    case 0x09:
      return "dashDot"
    case 0x0a:
      return "mediumDashDot"
    case 0x0b:
      return "dashDotDot"
    case 0x0c:
      return "mediumDashDotDot"
    case 0x0d:
      return "slantDashDot"
    default:
      return undefined
  }
}

export function biffFillPattern(code: number): FillPattern {
  switch (code) {
    case 0x02:
      return "mediumGray"
    case 0x03:
      return "darkGray"
    case 0x04:
      return "lightGray"
    case 0x05:
      return "darkHorizontal"
    case 0x06:
      return "darkVertical"
    case 0x07:
      return "darkDown"
    case 0x08:
      return "darkUp"
    case 0x09:
      return "darkGrid"
    case 0x0a:
      return "darkTrellis"
    case 0x0b:
      return "lightHorizontal"
    case 0x0c:
      return "lightVertical"
    case 0x0d:
      return "lightDown"
    case 0x0e:
      return "lightUp"
    case 0x0f:
      return "lightGrid"
    case 0x10:
      return "lightTrellis"
    case 0x11:
      return "gray125"
    case 0x12:
      return "gray0625"
    case 0x01:
      return "solid"
    case 0x00:
    default:
      return "none"
  }
}
