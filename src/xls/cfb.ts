// ── Compound File Binary Reader ──────────────────────────────────────
// Minimal MS-CFB/OLE2 reader used by the XLS BIFF reader. It supports
// regular FAT chains, mini streams, DIFAT extension sectors, and root-level
// stream lookup without pulling in a dependency.

import { ParseError } from "../errors"

const FREE_SECT = 0xffffffff
const END_OF_CHAIN = 0xfffffffe
const FAT_SECT = 0xfffffffd
const DIFAT_SECT = 0xfffffffc

const CFB_MAGIC = Object.freeze([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const)

interface DirectoryEntry {
  name: string
  type: number
  startSector: number
  size: number
}

/** Parsed stream entry exposed to callers. */
export interface CfbStreamEntry {
  name: string
  size: number
}

function concat(parts: Uint8Array[], totalLen: number): Uint8Array {
  const out = new Uint8Array(totalLen)
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function normalizeName(name: string): string {
  return name.toLowerCase().replace(/^\\+/, "")
}

function decodeUtf16Le(bytes: Uint8Array): string {
  // TextDecoder("utf-16le") is available in Node and all modern browsers, but
  // keeping a small fallback makes the parser usable in stricter edge runtimes.
  try {
    return new TextDecoder("utf-16le").decode(bytes)
  } catch {
    let s = ""
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      s += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8))
    }
    return s
  }
}

/**
 * Tiny Compound File Binary reader.
 *
 * XLS workbooks store their BIFF stream in an OLE2/CFB file. Password
 * protected OOXML files also use CFB, but they contain `EncryptedPackage`
 * instead of a BIFF `Workbook` stream; the XLS reader detects that before
 * attempting BIFF parsing.
 */
export class CfbReader {
  private readonly view: DataView
  private readonly sectorSize: number
  private readonly miniSectorSize: number
  private readonly miniStreamCutoff: number
  private readonly fat: number[]
  private readonly miniFat: number[]
  private readonly entries: DirectoryEntry[]
  private readonly streamMap = new Map<string, DirectoryEntry>()
  private readonly rootEntry?: DirectoryEntry
  private readonly miniStream: Uint8Array

  constructor(private readonly data: Uint8Array) {
    if (data.length < 512) {
      throw new ParseError("Invalid CFB file: file is smaller than the 512-byte header")
    }
    for (let i = 0; i < CFB_MAGIC.length; i++) {
      if (data[i] !== CFB_MAGIC[i]) {
        throw new ParseError("Invalid CFB file: missing OLE2 magic header")
      }
    }

    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength)

    const byteOrder = this.u16(0x1c)
    if (byteOrder !== 0xfffe) {
      throw new ParseError("Invalid CFB file: unsupported byte order")
    }

    const sectorShift = this.u16(0x1e)
    const miniSectorShift = this.u16(0x20)
    this.sectorSize = 1 << sectorShift
    this.miniSectorSize = 1 << miniSectorShift
    this.miniStreamCutoff = this.u32(0x38)

    if (this.sectorSize !== 512 && this.sectorSize !== 4096) {
      throw new ParseError(`Invalid CFB file: unsupported sector size ${this.sectorSize}`)
    }
    if (this.miniSectorSize !== 64) {
      throw new ParseError(`Invalid CFB file: unsupported mini sector size ${this.miniSectorSize}`)
    }

    const numFatSectors = this.u32(0x2c)
    const firstDirSector = this.u32(0x30)
    const firstMiniFatSector = this.u32(0x3c)
    const numMiniFatSectors = this.u32(0x40)
    const firstDifatSector = this.u32(0x44)
    const numDifatSectors = this.u32(0x48)

    const difat = this.readDifat(firstDifatSector, numDifatSectors, numFatSectors)
    this.fat = this.readFat(difat)

    const dirBytes = this.readRegularChain(firstDirSector)
    this.entries = this.readDirectory(dirBytes)
    this.rootEntry = this.entries.find((e) => e.type === 5)

    for (const entry of this.entries) {
      if (entry.type === 2) {
        this.streamMap.set(normalizeName(entry.name), entry)
      }
    }

    this.miniFat = this.readMiniFat(firstMiniFatSector, numMiniFatSectors)
    this.miniStream = this.rootEntry ? this.readRegularChain(this.rootEntry.startSector, this.rootEntry.size) : new Uint8Array(0)
  }

  /** Return all root-level streams discovered in the CFB directory. */
  listStreams(): CfbStreamEntry[] {
    return [...this.streamMap.values()].map((entry) => ({ name: entry.name, size: entry.size }))
  }

  /** Whether a stream exists. Name matching is case-insensitive. */
  hasStream(name: string): boolean {
    return this.streamMap.has(normalizeName(name))
  }

  /** Extract a stream by name. Name matching is case-insensitive. */
  getStream(name: string): Uint8Array | undefined {
    const entry = this.streamMap.get(normalizeName(name))
    if (!entry) return undefined
    if (entry.size === 0) return new Uint8Array(0)

    if (entry.size < this.miniStreamCutoff && entry.startSector !== END_OF_CHAIN) {
      return this.readMiniChain(entry.startSector, entry.size)
    }
    return this.readRegularChain(entry.startSector, entry.size)
  }

  private readDifat(firstDifatSector: number, numDifatSectors: number, numFatSectors: number): number[] {
    const difat: number[] = []

    for (let i = 0; i < 109; i++) {
      const sid = this.u32(0x4c + i * 4)
      if (sid !== FREE_SECT) difat.push(sid)
    }

    let next = firstDifatSector
    const entriesPerDifatSector = this.sectorSize / 4 - 1
    for (let i = 0; i < numDifatSectors && next !== END_OF_CHAIN && next !== FREE_SECT; i++) {
      const sector = this.readSector(next)
      const sectorView = new DataView(sector.buffer, sector.byteOffset, sector.byteLength)
      for (let j = 0; j < entriesPerDifatSector; j++) {
        const sid = sectorView.getUint32(j * 4, true)
        if (sid !== FREE_SECT) difat.push(sid)
      }
      next = sectorView.getUint32(this.sectorSize - 4, true)
    }

    if (difat.length < numFatSectors) {
      throw new ParseError("Invalid CFB file: DIFAT does not declare all FAT sectors")
    }

    return difat.slice(0, numFatSectors)
  }

  private readFat(difat: number[]): number[] {
    const fat: number[] = []
    for (const fatSector of difat) {
      if (fatSector === FAT_SECT || fatSector === DIFAT_SECT || fatSector === END_OF_CHAIN) continue
      const sector = this.readSector(fatSector)
      const sectorView = new DataView(sector.buffer, sector.byteOffset, sector.byteLength)
      for (let pos = 0; pos < this.sectorSize; pos += 4) {
        fat.push(sectorView.getUint32(pos, true))
      }
    }
    return fat
  }

  private readMiniFat(firstMiniFatSector: number, numMiniFatSectors: number): number[] {
    if (firstMiniFatSector === END_OF_CHAIN || firstMiniFatSector === FREE_SECT || numMiniFatSectors === 0) {
      return []
    }
    const bytes = this.readRegularChain(firstMiniFatSector, numMiniFatSectors * this.sectorSize)
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const miniFat: number[] = []
    for (let pos = 0; pos + 4 <= bytes.length; pos += 4) {
      miniFat.push(view.getUint32(pos, true))
    }
    return miniFat
  }

  private readDirectory(dirBytes: Uint8Array): DirectoryEntry[] {
    const entries: DirectoryEntry[] = []
    const view = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength)

    for (let pos = 0; pos + 128 <= dirBytes.length; pos += 128) {
      const nameLen = view.getUint16(pos + 64, true)
      const type = dirBytes[pos + 66] ?? 0
      if (nameLen < 2 || type === 0) continue

      const nameBytes = dirBytes.subarray(pos, pos + Math.min(nameLen - 2, 64))
      const name = decodeUtf16Le(nameBytes)
      const startSector = view.getUint32(pos + 116, true)
      const sizeLow = view.getUint32(pos + 120, true)
      const sizeHigh = view.getUint32(pos + 124, true)
      const size = sizeHigh * 0x1_0000_0000 + sizeLow

      entries.push({ name, type, startSector, size })
    }

    return entries
  }

  private readRegularChain(startSector: number, maxBytes?: number): Uint8Array {
    if (startSector === END_OF_CHAIN || startSector === FREE_SECT) return new Uint8Array(0)

    const parts: Uint8Array[] = []
    const seen = new Set<number>()
    let totalLen = 0
    let sector = startSector

    while (sector !== END_OF_CHAIN && sector !== FREE_SECT) {
      if (sector === FAT_SECT || sector === DIFAT_SECT) {
        throw new ParseError("Invalid CFB file: stream chain points at a FAT/DIFAT sector")
      }
      if (sector >= this.fat.length) {
        throw new ParseError("Invalid CFB file: sector chain points outside the FAT")
      }
      if (seen.has(sector)) {
        throw new ParseError("Invalid CFB file: cyclic sector chain")
      }
      seen.add(sector)

      const chunk = this.readSector(sector)
      parts.push(chunk)
      totalLen += chunk.length

      if (maxBytes !== undefined && totalLen >= maxBytes) break
      sector = this.fat[sector]!
    }

    const out = concat(parts, totalLen)
    return maxBytes === undefined ? out : out.subarray(0, Math.min(maxBytes, out.length))
  }

  private readMiniChain(startMiniSector: number, maxBytes: number): Uint8Array {
    const parts: Uint8Array[] = []
    const seen = new Set<number>()
    let totalLen = 0
    let miniSector = startMiniSector

    while (miniSector !== END_OF_CHAIN && miniSector !== FREE_SECT) {
      if (miniSector >= this.miniFat.length) {
        throw new ParseError("Invalid CFB file: mini stream chain points outside the mini FAT")
      }
      if (seen.has(miniSector)) {
        throw new ParseError("Invalid CFB file: cyclic mini sector chain")
      }
      seen.add(miniSector)

      const offset = miniSector * this.miniSectorSize
      if (offset + this.miniSectorSize > this.miniStream.length) {
        throw new ParseError("Invalid CFB file: mini sector points outside the root mini stream")
      }
      const chunk = this.miniStream.subarray(offset, offset + this.miniSectorSize)
      parts.push(chunk)
      totalLen += chunk.length

      if (totalLen >= maxBytes) break
      miniSector = this.miniFat[miniSector]!
    }

    return concat(parts, totalLen).subarray(0, maxBytes)
  }

  private readSector(sector: number): Uint8Array {
    const offset = (sector + 1) * this.sectorSize
    if (offset < 0 || offset + this.sectorSize > this.data.length) {
      throw new ParseError("Invalid CFB file: sector offset outside file")
    }
    return this.data.subarray(offset, offset + this.sectorSize)
  }

  private u16(offset: number): number {
    return this.view.getUint16(offset, true)
  }

  private u32(offset: number): number {
    return this.view.getUint32(offset, true)
  }
}
