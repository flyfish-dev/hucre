import type { SheetImage } from "../_types"
import type { XmlElement } from "../xml/parser"

type Layout = Pick<SheetImage["anchor"], "kind" | "editAs" | "extent" | "position">

function child(parent: XmlElement | undefined, name: string): XmlElement | undefined {
  return parent?.children.find(
    (node): node is XmlElement => typeof node !== "string" && (node.local || node.tag) === name,
  )
}
function pair(node: XmlElement | undefined, a: string, b: string): [number, number] | undefined {
  if (!node || node.attrs[a] === undefined || node.attrs[b] === undefined) return undefined
  if (!/^[+-]?\d+$/.test(node.attrs[a].trim()) || !/^[+-]?\d+$/.test(node.attrs[b].trim()))
    return undefined
  const x = Number(node.attrs[a]),
    y = Number(node.attrs[b])
  return Number.isSafeInteger(x) && Number.isSafeInteger(y) ? [x, y] : undefined
}

/**
 * Container semantics belong to the parser, not a renderer's second ZIP read.
 * Keep EMUs intact: rounding to whole 96-DPI pixels here loses precision at
 * other DPIs/zoom levels. A twoCellAnchor with editAs=oneCell still has two
 * saved markers; consumers need both the markers AND its resizing behavior.
 */
export function readDrawingLayout(anchor: XmlElement): Layout {
  const local = anchor.local || anchor.tag
  const kind =
    local === "twoCellAnchor"
      ? "twoCell"
      : local === "oneCellAnchor"
        ? "oneCell"
        : local === "absoluteAnchor"
          ? "absolute"
          : undefined
  if (!kind) return {}
  const result: Layout = { kind }
  const editAs = anchor.attrs["editAs"]
  if (
    kind === "twoCell" &&
    (editAs === "oneCell" || editAs === "absolute" || editAs === "twoCell")
  ) {
    result.editAs = editAs
  }
  const xfrm = child(child(child(anchor, "pic"), "spPr"), "xfrm")
  const ext = kind === "twoCell" ? child(xfrm, "ext") : child(anchor, "ext")
  const size = pair(ext, "cx", "cy")
  if (size && size[0] >= 0 && size[1] >= 0) result.extent = { cx: size[0], cy: size[1] }
  const pos = pair(
    kind === "absolute"
      ? child(anchor, "pos")
      : editAs === "absolute"
        ? child(xfrm, "off")
        : undefined,
    "x",
    "y",
  )
  if (pos) result.position = { x: pos[0], y: pos[1] }
  return result
}
