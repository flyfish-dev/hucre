import type { SheetShapeColor, SheetShapePrimitive } from "../_types"
import type { XmlElement } from "../xml/parser"

type Mapper = { x: number; y: number; sx: number; sy: number }
const MAX_GROUP_DEPTH = 32
const MAX_PRIMITIVES_PER_ANCHOR = 4096

function child(parent: XmlElement | undefined, name: string): XmlElement | undefined {
  return parent?.children.find(
    (node): node is XmlElement => typeof node !== "string" && (node.local || node.tag) === name,
  )
}

function pair(node: XmlElement | undefined, a: string, b: string): [number, number] | undefined {
  if (!node || node.attrs[a] === undefined || node.attrs[b] === undefined) return undefined
  if (!/^[+-]?\d+$/.test(node.attrs[a].trim()) || !/^[+-]?\d+$/.test(node.attrs[b].trim())) {
    return undefined
  }
  const x = Number(node.attrs[a])
  const y = Number(node.attrs[b])
  return Number.isSafeInteger(x) && Number.isSafeInteger(y) ? [x, y] : undefined
}

function color(node: XmlElement | undefined): SheetShapeColor | null | undefined {
  if (!node) return undefined
  if (child(node, "noFill")) return null
  const solid = child(node, "solidFill")
  if (!solid) return undefined
  const rgb = child(solid, "srgbClr")?.attrs["val"] ?? child(solid, "sysClr")?.attrs["lastClr"]
  if (rgb && /^[0-9a-f]{6}$/i.test(rgb)) return { rgb: rgb.toUpperCase() }
  const scheme = child(solid, "schemeClr")?.attrs["val"]
  if (scheme && /^[a-z][a-z0-9]*$/i.test(scheme)) return { scheme }
  return undefined
}

function primitive(
  shape: XmlElement,
  rect: { x: number; y: number; width: number; height: number },
): SheetShapePrimitive | undefined {
  const properties = child(shape, "spPr")
  const geometry = child(properties, "prstGeom")?.attrs["prst"]
  if (geometry !== "rect" && geometry !== "ellipse") return undefined
  if (
    ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) ||
    rect.width <= 0 ||
    rect.height <= 0
  )
    return undefined
  const result: SheetShapePrimitive = { geometry, ...rect }
  const fill = color(properties)
  if (fill !== undefined) result.fill = fill
  const line = child(properties, "ln")
  const stroke = color(line)
  if (stroke !== undefined) result.stroke = stroke
  if (line) {
    const width = Number(line.attrs["w"])
    if (Number.isSafeInteger(width) && width >= 0) result.strokeWidth = width / 9_525
  }
  return result
}

function appendGroup(
  group: XmlElement,
  map: Mapper,
  result: SheetShapePrimitive[],
  depth: number,
): void {
  if (depth >= MAX_GROUP_DEPTH || result.length >= MAX_PRIMITIVES_PER_ANCHOR) return
  for (const node of group.children) {
    if (result.length >= MAX_PRIMITIVES_PER_ANCHOR) break
    if (typeof node === "string") continue
    const local = node.local || node.tag
    if (local !== "sp" && local !== "grpSp") continue
    const properties = child(node, local === "sp" ? "spPr" : "grpSpPr")
    const transform = child(properties, "xfrm")
    const offset = pair(child(transform, "off"), "x", "y")
    const extent = pair(child(transform, "ext"), "cx", "cy")
    if (!offset || !extent || extent[0] <= 0 || extent[1] <= 0) continue
    const rect = {
      x: map.x + offset[0] * map.sx,
      y: map.y + offset[1] * map.sy,
      width: extent[0] * map.sx,
      height: extent[1] * map.sy,
    }
    if (local === "sp") {
      const shape = primitive(node, rect)
      if (shape) result.push(shape)
      continue
    }
    const childOffset = pair(child(transform, "chOff"), "x", "y")
    const childExtent = pair(child(transform, "chExt"), "cx", "cy")
    if (!childOffset || !childExtent || childExtent[0] <= 0 || childExtent[1] <= 0) continue
    const sx = rect.width / childExtent[0]
    const sy = rect.height / childExtent[1]
    appendGroup(
      node,
      {
        x: rect.x - childOffset[0] * sx,
        y: rect.y - childOffset[1] * sy,
        sx,
        sy,
      },
      result,
      depth + 1,
    )
  }
}

/** Extract basic vector geometry without treating grouped coordinates as sheet coordinates. */
export function readDrawingShapePrimitives(anchor: XmlElement): SheetShapePrimitive[] {
  const direct = child(anchor, "sp")
  if (direct) {
    const shape = primitive(direct, { x: 0, y: 0, width: 1, height: 1 })
    return shape ? [shape] : []
  }
  const group = child(anchor, "grpSp")
  if (!group) return []
  const transform = child(child(group, "grpSpPr"), "xfrm")
  const offset = pair(child(transform, "chOff"), "x", "y")
  const extent = pair(child(transform, "chExt"), "cx", "cy")
  if (!offset || !extent || extent[0] <= 0 || extent[1] <= 0) return []
  const result: SheetShapePrimitive[] = []
  appendGroup(
    group,
    {
      x: -offset[0] / extent[0],
      y: -offset[1] / extent[1],
      sx: 1 / extent[0],
      sy: 1 / extent[1],
    },
    result,
    0,
  )
  return result
}
