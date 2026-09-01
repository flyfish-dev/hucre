// ── BIFF8 Record Stream ──────────────────────────────────────────────
// The "Workbook" stream inside an .xls OLE2 container is a flat sequence
// of BIFF records: 2-byte record id + 2-byte length + `length` bytes of
// data (each record body is ≤ 8224 bytes; longer data overflows into
// CONTINUE records). See [MS-XLS].

import { ParseError } from "../errors"

export interface BiffRecord {
  /** Record id (sid). */
  id: number
  data: Uint8Array
  /** Byte offset of this record's header within the stream. */
  offset: number
}

export const SID = {
  FORMULA: 0x0006,
  EOF: 0x000a,
  HORIZONTALPAGEBREAKS: 0x001b,
  VERTICALPAGEBREAKS: 0x001a,
  LEFTMARGIN: 0x0026,
  RIGHTMARGIN: 0x0027,
  TOPMARGIN: 0x0028,
  BOTTOMMARGIN: 0x0029,
  PRINTHEADERS: 0x002a,
  PRINTGRIDLINES: 0x002b,
  CONTINUE: 0x003c,
  FONT: 0x0031,
  DEFCOLWIDTH: 0x0055,
  COLINFO: 0x007d,
  WSBOOL: 0x0081,
  HCENTER: 0x0083,
  VCENTER: 0x0084,
  DATEMODE: 0x0022,
  PALETTE: 0x0092,
  STANDARDWIDTH: 0x0099,
  SCL: 0x00a0,
  SETUP: 0x00a1,
  BLANK: 0x0201,
  DIMENSIONS: 0x0200,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  STRING: 0x0207,
  ROW: 0x0208,
  WINDOW2: 0x023e,
  INDEX: 0x020b,
  RK: 0x027e,
  MULRK: 0x00bd,
  MULBLANK: 0x00be,
  LABELSST: 0x00fd,
  // RSTRING is a BIFF5/BIFF7 record (rich-text cell: rw, col, ixfe, a
  // codepage byte string, then a run count and its runs). BIFF8 dropped it:
  // rich text moved into the SST, so those cells arrive as LABELSST. The
  // reader rejects anything that is not BIFF8 (see the version gate in
  // reader.ts), so this sid cannot reach a cell handler — listed for
  // recognition only, deliberately not parsed. See #411.
  RSTRING: 0x00d6,
  SST: 0x00fc,
  XF: 0x00e0,
  FORMAT: 0x041e,
  BOUNDSHEET: 0x0085,
  MERGECELLS: 0x00e5,
  DEFAULTROWHEIGHT: 0x0225,
  BOF: 0x0809,
} as const

/** Parse a stream into records, recording each one's byte offset. */
export function parseRecords(stream: Uint8Array): BiffRecord[] {
  const out: BiffRecord[] = []
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
  let pos = 0
  while (pos + 4 <= stream.length) {
    const id = view.getUint16(pos, true)
    const len = view.getUint16(pos + 2, true)
    const start = pos
    pos += 4
    out.push({ id, data: stream.subarray(pos, pos + len), offset: start })
    pos += len
    // A zero id past real data means padding — stop.
    if (id === 0 && len === 0) break
  }
  return out
}

/** Little-endian cursor over a record body. */
export class Reader {
  pos = 0
  private view: DataView

  constructor(public buf: Uint8Array) {
    this.view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  }

  u8(): number {
    return this.buf[this.pos++]
  }
  u16(): number {
    const v = this.view.getUint16(this.pos, true)
    this.pos += 2
    return v
  }
  i16(): number {
    const v = this.view.getInt16(this.pos, true)
    this.pos += 2
    return v
  }
  u32(): number {
    const v = this.view.getUint32(this.pos, true)
    this.pos += 4
    return v
  }
  f64(): number {
    const v = this.view.getFloat64(this.pos, true)
    this.pos += 8
    return v
  }
  skip(n: number): void {
    this.pos += n
  }
  remaining(): number {
    return this.buf.length - this.pos
  }
}

/** Decode an RK number (MS-XLS §2.5.166). */
export function decodeRk(rk: number): number {
  const fX100 = (rk & 1) !== 0
  const fInt = (rk & 2) !== 0
  let value: number
  if (fInt) {
    value = (rk | 0) >> 2
  } else {
    const buf = new ArrayBuffer(8)
    const dv = new DataView(buf)
    dv.setUint32(4, rk & 0xfffffffc, true)
    value = dv.getFloat64(0, true)
  }
  return fX100 ? value / 100 : value
}

// ── SST (shared string table) with CONTINUE handling ─────────────────
// Strings split across CONTINUE records on character boundaries; each
// CONTINUE that resumes a string's character array restarts with a 1-byte
// option flag (fHighByte). Header fields and rich/phonetic trailers are
// assumed not to straddle a boundary (how Excel writes them).

/** Reads across the SST record + its trailing CONTINUE blocks. */
class BlockStream {
  private bi = 0
  private pos = 0
  constructor(private blocks: Uint8Array[]) {}

  private cur(): Uint8Array {
    return this.blocks[this.bi]
  }
  remainingInBlock(): number {
    const b = this.cur()
    return b ? b.length - this.pos : 0
  }
  atEnd(): boolean {
    return this.bi >= this.blocks.length
  }
  /** Total bytes across all blocks — an upper bound on decodable records. */
  totalBytes(): number {
    let n = 0
    for (const b of this.blocks) n += b.length
    return n
  }
  /** Move to the next block (used when a string's chars continue). */
  nextBlock(): void {
    this.bi++
    this.pos = 0
  }
  ensure(): void {
    while (!this.atEnd() && this.remainingInBlock() === 0) this.nextBlock()
  }
  u8(): number {
    const block = this.blocks[this.bi]
    if (block && this.pos < block.length) return block[this.pos++]
    this.ensure()
    if (this.atEnd()) throw new ParseError("Invalid XLS: truncated shared string table")
    return this.cur()[this.pos++]
  }
  u16(): number {
    const block = this.blocks[this.bi]
    if (block && this.pos + 2 <= block.length) {
      const value = block[this.pos] | (block[this.pos + 1] << 8)
      this.pos += 2
      return value
    }
    return this.u8() | (this.u8() << 8)
  }
  u32(): number {
    const block = this.blocks[this.bi]
    if (block && this.pos + 4 <= block.length) {
      const value =
        (block[this.pos] |
          (block[this.pos + 1] << 8) |
          (block[this.pos + 2] << 16) |
          (block[this.pos + 3] << 24)) >>>
        0
      this.pos += 4
      return value
    }
    return (this.u8() | (this.u8() << 8) | (this.u8() << 16) | (this.u8() << 24)) >>> 0
  }
  skip(n: number): void {
    // skip may cross blocks (rich/phonetic trailers), no grbit byte
    let left = n
    while (left > 0) {
      this.ensure()
      const take = Math.min(left, this.remainingInBlock())
      // Once the blocks are exhausted, remainingInBlock() returns 0 and
      // `left` would never decrease — a live lock no try/catch can
      // interrupt. The callers feed this untrusted lengths (`cRun * 4`
      // from a u16, `cbExt` from a u32), so a file claiming more
      // trailer bytes than it carries used to hang the process. See #389.
      if (take === 0) {
        throw new ParseError(
          `Invalid XLS: string record claims ${n} trailing bytes but the stream ends after ${n - left}`,
        )
      }
      this.pos += take
      left -= take
    }
  }
}

export interface BiffSstRun {
  start: number
  fontIndex: number
}

export interface BiffSstEntry {
  text: string
  runs?: BiffSstRun[]
}

/**
 * Parse an SST record (plus its following CONTINUE records) into the
 * shared-string array. `blocks` is `[sstData, ...continueDatas]`.
 */
export function parseSst(blocks: Uint8Array[]): string[] {
  return parseSstEntries(blocks).map((entry) => entry.text)
}

/** Parse SST text together with BIFF8 rich-text formatting runs. */
export function parseSstEntries(blocks: Uint8Array[]): BiffSstEntry[] {
  const s = new BlockStream(blocks)
  s.skip(4) // cstTotal
  const cstUnique = s.u32()
  // cstUnique is an untrusted u32. Each string costs at least 3 bytes
  // (2-byte cch + 1-byte grbit), so a string count larger than the total
  // available bytes is a corrupt / hostile header — cap the loop by the
  // bytes actually present so we don't spin allocating empty strings.
  const maxStrings = Math.min(cstUnique, s.totalBytes())
  const out: BiffSstEntry[] = []
  for (let i = 0; i < maxStrings; i++) {
    s.ensure()
    if (s.atEnd()) break
    out.push(readSstString(s))
  }
  return out
}

function readSstString(s: BlockStream): BiffSstEntry {
  const cch = s.u16()
  let grbit = s.u8()
  let compressed = (grbit & 0x01) === 0
  const rich = (grbit & 0x08) !== 0
  const phonetic = (grbit & 0x04) !== 0
  const cRun = rich ? s.u16() : 0
  const cbExt = phonetic ? s.u32() : 0

  let str = ""
  let read = 0
  while (read < cch) {
    if (s.remainingInBlock() === 0) {
      // Continue into the next block: it restarts with an option flag for
      // the remaining characters.
      s.nextBlock()
      grbit = s.u8()
      compressed = (grbit & 0x01) === 0
    }
    str += String.fromCharCode(compressed ? s.u8() : s.u16())
    read++
  }
  const runs: BiffSstRun[] = []
  for (let index = 0; index < cRun; index++) {
    runs.push({ start: s.u16(), fontIndex: s.u16() })
  }
  if (phonetic) s.skip(cbExt)
  return runs.length > 0 ? { text: str, runs } : { text: str }
}
