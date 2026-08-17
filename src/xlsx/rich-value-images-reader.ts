// ── Excel 365 In-Cell Picture Reader ────────────────────────────────
//
// Excel's "Place in Cell" pictures do not use WPS' cellimages.xml /
// DISPIMG representation. They are rich values wired through four parts:
//
//   xl/metadata.xml                     cell @vm -> rich-value record
//   xl/richData/rdrichvaluestructure.xml record field definitions
//   xl/richData/rdrichvalue.xml          field values
//   xl/richData/richValueRel.xml         field index -> image rId
//
// The final rId resolves through richValueRel.xml.rels to xl/media/*.

import { parseXml } from "../xml/parser"
import type { XmlElement } from "../xml/parser"
import { childElements, findChild } from "../xml/tree"

const XLRICHVALUE = "XLRICHVALUE"
const LOCAL_IMAGE_KEY = "_rvRel:LocalImageIdentifier"
const IMAGE_ID_PREFIX = "excel365-rich-value-"

export interface ExcelRichValueImageRef {
  /** hucre workbook-level id assigned to the rich-value record. */
  id: string
  /** Zero-based rich-value record index. */
  richValueIndex: number
  /** rId in richValueRel.xml.rels. */
  embedRId: string
}

export interface ParsedExcelRichValueImages {
  /** Image-bearing rich-value records in workbook order. */
  refs: ExcelRichValueImageRef[]
  /** One-based worksheet `vm` index -> CellImage id. */
  imageIdByVm: Map<number, string>
}

/**
 * Resolve Excel 365 XLRICHVALUE metadata and rich-data records into image
 * references. The caller owns ZIP path resolution and binary extraction.
 */
export function parseExcelRichValueImages(
  metadataXml: string,
  structuresXml: string,
  valuesXml: string,
  richValueRelsXml: string,
): ParsedExcelRichValueImages {
  const richValueByVm = parseRichValueMetadata(metadataXml)
  const structures = parseStructures(structuresXml)
  const relationshipIds = parseRichValueRelationshipIds(richValueRelsXml)
  const refs = parseImageRefs(valuesXml, structures, relationshipIds)
  const idByRichValue = new Map(refs.map((ref) => [ref.richValueIndex, ref.id]))
  const imageIdByVm = new Map<number, string>()

  for (const [vm, richValueIndex] of richValueByVm) {
    const id = idByRichValue.get(richValueIndex)
    if (id) imageIdByVm.set(vm, id)
  }

  return { refs, imageIdByVm }
}

/** Map one-based worksheet `vm` indexes to zero-based rich-value records. */
function parseRichValueMetadata(xml: string): Map<number, number> {
  const root = parseXml(xml)
  const metadata = root.local === "metadata" ? root : findChild(root, "metadata")
  const out = new Map<number, number>()
  if (!metadata) return out

  const typeNames: string[] = []
  const metadataTypes = findChild(metadata, "metadataTypes")
  if (metadataTypes) {
    for (const type of childElements(metadataTypes)) {
      if (type.local === "metadataType") typeNames.push(type.attrs.name ?? "")
    }
  }

  const richValueBlocks: Array<number | undefined> = []
  for (const future of childElements(metadata)) {
    if (future.local !== "futureMetadata" || future.attrs.name !== XLRICHVALUE) continue
    for (const block of childElements(future)) {
      if (block.local !== "bk") continue
      const marker = findDescendant(block, "rvb")
      richValueBlocks.push(nonNegativeInteger(marker?.attrs.i))
    }
  }

  const valueMetadata = findChild(metadata, "valueMetadata")
  if (!valueMetadata) return out

  let vm = 0
  for (const block of childElements(valueMetadata)) {
    if (block.local !== "bk") continue
    vm++
    for (const record of childElements(block)) {
      if (record.local !== "rc") continue
      const typeIndex = positiveInteger(record.attrs.t)
      const blockIndex = nonNegativeInteger(record.attrs.v)
      if (!typeIndex || blockIndex === undefined) continue
      if (typeNames[typeIndex - 1] !== XLRICHVALUE) continue
      const richValueIndex = richValueBlocks[blockIndex]
      if (richValueIndex !== undefined) out.set(vm, richValueIndex)
    }
  }

  return out
}

/** Read the ordered key list for every rich-value structure. */
function parseStructures(xml: string): string[][] {
  const root = parseXml(xml)
  const structuresRoot = root.local === "rvStructures" ? root : findChild(root, "rvStructures")
  if (!structuresRoot) return []

  const structures: string[][] = []
  for (const structure of childElements(structuresRoot)) {
    if (structure.local !== "s") continue
    structures.push(
      childElements(structure)
        .filter((key) => key.local === "k")
        .map((key) => key.attrs.n ?? ""),
    )
  }
  return structures
}

/** Read the ordered rIds from richValueRel.xml. */
function parseRichValueRelationshipIds(xml: string): string[] {
  const root = parseXml(xml)
  const relationshipsRoot = root.local === "richValueRels" ? root : findChild(root, "richValueRels")
  if (!relationshipsRoot) return []

  const ids: string[] = []
  for (const rel of childElements(relationshipsRoot)) {
    if (rel.local !== "rel") continue
    const id = namespacedAttribute(rel.attrs, "id")
    if (id) ids.push(id)
  }
  return ids
}

function parseImageRefs(
  xml: string,
  structures: readonly string[][],
  relationshipIds: readonly string[],
): ExcelRichValueImageRef[] {
  const root = parseXml(xml)
  const dataRoot = root.local === "rvData" ? root : findChild(root, "rvData")
  if (!dataRoot) return []

  const refs: ExcelRichValueImageRef[] = []
  let richValueIndex = 0
  for (const richValue of childElements(dataRoot)) {
    if (richValue.local !== "rv") continue
    const structureIndex = nonNegativeInteger(richValue.attrs.s)
    const keys = structureIndex === undefined ? undefined : structures[structureIndex]
    const imageField = keys?.indexOf(LOCAL_IMAGE_KEY) ?? -1
    if (imageField >= 0) {
      const values = childElements(richValue).filter((value) => value.local === "v")
      const relationshipIndex = nonNegativeInteger(textContent(values[imageField]))
      const embedRId =
        relationshipIndex === undefined ? undefined : relationshipIds[relationshipIndex]
      if (embedRId) {
        refs.push({
          id: `${IMAGE_ID_PREFIX}${richValueIndex}`,
          richValueIndex,
          embedRId,
        })
      }
    }
    richValueIndex++
  }
  return refs
}

function findDescendant(el: XmlElement, localName: string): XmlElement | undefined {
  for (const child of childElements(el)) {
    if (child.local === localName) return child
    const nested = findDescendant(child, localName)
    if (nested) return nested
  }
  return undefined
}

function namespacedAttribute(attrs: Record<string, string>, localName: string): string | undefined {
  if (attrs[localName]) return attrs[localName]
  for (const [name, value] of Object.entries(attrs)) {
    if (name.endsWith(`:${localName}`)) return value
  }
  return undefined
}

function textContent(el: XmlElement | undefined): string | undefined {
  if (!el) return undefined
  return el.children.filter((child): child is string => typeof child === "string").join("")
}

function nonNegativeInteger(value: string | undefined): number | undefined {
  if (value === undefined || value.trim() === "") return undefined
  const parsed = Number(value)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined
}

function positiveInteger(value: string | undefined): number | undefined {
  const parsed = nonNegativeInteger(value)
  return parsed !== undefined && parsed > 0 ? parsed : undefined
}
