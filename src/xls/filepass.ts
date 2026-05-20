// ── XLS FilePass Decryption Helpers ──────────────────────────────────
// Minimal MS-XLS/MS-OFFCRYPTO legacy BIFF encryption support.
//
// BIFF workbooks can encrypt the Workbook stream at the record-payload level
// using one of three families carried by the FilePass record:
// - XOR obfuscation (very old workbooks)
// - RC4 Standard
// - RC4 CryptoAPI
//
// The code below intentionally has no npm dependency: RC4, MD5 and SHA-1 are
// implemented in-place so the synchronous BIFF reader can decrypt records as
// they are read. Modern encrypted OOXML packages are handled separately by
// src/crypto/office-crypto.ts.

import { EncryptedFileError, ParseError } from "../errors"

const FILEPASS_XOR = 0x0000
const FILEPASS_RC4 = 0x0001
const XOR_ARRAY_LEN = 16
const RC4_BLOCK_SIZE = 1024

export interface BiffRecordDecryptor {
  readonly method: "xor" | "rc4" | "cryptoapi"
  decryptRecordPayload(recordOffset: number, sid: number, data: Uint8Array): Uint8Array
}

interface Rc4FilePassInfo {
  method: "rc4" | "cryptoapi"
  salt: Uint8Array
  encryptedVerifier: Uint8Array
  encryptedVerifierHash: Uint8Array
  encryptedVerifierHashSize: number
  keySize: number
  hash: "md5" | "sha1"
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function u16(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint16(offset, true)
}

function u32(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint32(offset, true)
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, p) => sum + p.length, 0)
  const out = new Uint8Array(total)
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function utf16le(password: string): Uint8Array {
  const out = new Uint8Array(password.length * 2)
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i)
    out[i * 2] = c & 0xff
    out[i * 2 + 1] = c >>> 8
  }
  return out
}

function int32le(n: number): Uint8Array {
  const out = new Uint8Array(4)
  view(out).setUint32(0, n >>> 0, true)
  return out
}

function constantTimeStartsWith(actual: Uint8Array, expected: Uint8Array): boolean {
  let diff = actual.length < expected.length ? 1 : 0
  for (let i = 0; i < expected.length; i++) diff |= (actual[i] ?? 0) ^ (expected[i] ?? 0)
  return diff === 0
}

function clampPassword(password: string): string {
  // Legacy binary Office encryption uses the first 255 UTF-16 code units.
  return password.slice(0, 255)
}

export function createBiffFilePassDecryptor(
  filePass: Uint8Array,
  password: string | undefined,
): BiffRecordDecryptor {
  if (!password) {
    throw new EncryptedFileError(
      "xls",
      'XLS workbook is encrypted with a BIFF FilePass record. Pass `{ password: "..." }` in read options to decrypt it.',
    )
  }
  if (filePass.length < 2) throw new ParseError("Invalid XLS FilePass record: too short")

  const encryptionType = u16(filePass, 0)
  if (encryptionType === FILEPASS_XOR) return createXorDecryptor(filePass.subarray(2), password)
  if (encryptionType === FILEPASS_RC4) return createRc4Decryptor(filePass.subarray(2), password)

  throw new EncryptedFileError(
    "xls",
    `Unsupported XLS FilePass encryption type 0x${encryptionType.toString(16)}`,
  )
}

function createXorDecryptor(data: Uint8Array, password: string): BiffRecordDecryptor {
  if (data.length < 4) throw new ParseError("Invalid XOR FilePass record: missing key/verifier")
  const key = u16(data, 0)
  const verifier = u16(data, 2)
  const normalized = clampPassword(password)
  if (createXorVerifier(normalized) !== verifier) {
    throw new EncryptedFileError("xls", "Incorrect password for XOR-encrypted XLS workbook.")
  }
  const xorArray = createXorArray(normalized, key)

  return {
    method: "xor",
    decryptRecordPayload(recordOffset, _sid, bytes) {
      const out = new Uint8Array(bytes.length)
      for (let i = 0; i < bytes.length; i++) {
        const keyByte = xorArray[(recordOffset + 4 + i) % XOR_ARRAY_LEN]!
        // MS-XLS XOR obfuscation rotates encrypted bytes before applying the
        // derived XOR array. This path covers old Excel BIFF FilePass records;
        // modern RC4/CryptoAPI records use the stream decryptor below.
        out[i] = rotateRight8(bytes[i]! ^ keyByte, 3)
      }
      return out
    },
  }
}

function createXorVerifier(password: string): number {
  let verifier = 0
  for (let i = password.length - 1; i >= 0; i--) {
    verifier = rotateLeft15(verifier)
    verifier ^= password.charCodeAt(i) & 0x7f
  }
  verifier = rotateLeft15(verifier)
  verifier ^= password.length
  verifier ^= 0xce4b
  return verifier & 0xffff
}

function createXorArray(password: string, key: number): Uint8Array {
  const array = new Uint8Array(XOR_ARRAY_LEN)
  const seed = new Uint8Array(XOR_ARRAY_LEN)
  for (let i = 0; i < XOR_ARRAY_LEN; i++) seed[i] = 0xbb
  const p = clampPassword(password)
  for (let i = 0; i < Math.min(p.length, XOR_ARRAY_LEN); i++) {
    seed[i] = p.charCodeAt(p.length - 1 - i) & 0xff
  }
  for (let i = 0; i < XOR_ARRAY_LEN; i++) {
    const k = i % 2 === 0 ? key & 0xff : key >>> 8
    array[i] = rotateLeft8(seed[i]! ^ k, i & 0x07)
  }
  return array
}

function rotateLeft15(value: number): number {
  return (((value << 1) & 0x7fff) | ((value >>> 14) & 0x0001)) & 0x7fff
}

function rotateLeft8(value: number, bits: number): number {
  return ((value << bits) | (value >>> (8 - bits))) & 0xff
}

function rotateRight8(value: number, bits: number): number {
  return ((value >>> bits) | (value << (8 - bits))) & 0xff
}

function createRc4Decryptor(data: Uint8Array, password: string): BiffRecordDecryptor {
  const info = parseRc4FilePass(data)
  const normalized = clampPassword(password)
  const baseHash =
    info.method === "cryptoapi"
      ? digest(info.hash, concat([info.salt, utf16le(normalized)]))
      : md5(concat([utf16le(normalized), info.salt]))

  const verifierKey = deriveRc4BlockKey(baseHash, 0, info)
  const verifier = rc4(verifierKey, info.encryptedVerifier)
  const verifierHash = rc4(verifierKey, info.encryptedVerifierHash)
  const expectedHash = digest(info.hash, verifier).subarray(0, info.encryptedVerifierHashSize)
  if (!constantTimeStartsWith(verifierHash, expectedHash)) {
    throw new EncryptedFileError("xls", "Incorrect password for RC4-encrypted XLS workbook.")
  }

  const keyCache = new Map<number, Uint8Array>()
  const keyForBlock = (block: number): Uint8Array => {
    const cached = keyCache.get(block)
    if (cached) return cached
    const key = deriveRc4BlockKey(baseHash, block, info)
    keyCache.set(block, key)
    return key
  }

  return {
    method: info.method,
    decryptRecordPayload(recordOffset, _sid, bytes) {
      const out = new Uint8Array(bytes.length)
      let done = 0
      let absolute = recordOffset + 4
      while (done < bytes.length) {
        const block = Math.floor(absolute / RC4_BLOCK_SIZE)
        const blockOffset = absolute % RC4_BLOCK_SIZE
        const take = Math.min(bytes.length - done, RC4_BLOCK_SIZE - blockOffset)
        const encrypted = bytes.subarray(done, done + take)
        const keystream = rc4Drop(keyForBlock(block), blockOffset, take)
        for (let i = 0; i < take; i++) out[done + i] = encrypted[i]! ^ keystream[i]!
        done += take
        absolute += take
      }
      return out
    },
  }
}

function parseRc4FilePass(data: Uint8Array): Rc4FilePassInfo {
  // BIFF8 RC4 Standard verifier: salt(16), encryptedVerifier(16), encryptedVerifierHash(16)
  if (data.length >= 48 && data.length < 60) {
    return {
      method: "rc4",
      salt: data.subarray(0, 16),
      encryptedVerifier: data.subarray(16, 32),
      encryptedVerifierHash: data.subarray(32, 48),
      encryptedVerifierHashSize: 16,
      keySize: 40,
      hash: "md5",
    }
  }

  // CryptoAPI EncryptionHeader + EncryptionVerifier. The header starts with a
  // flags dword and size dword; the verifier immediately follows the variable
  // header. We read the common fields used by Excel-generated XLS files and
  // tolerate overlong provider names by trusting EncryptionHeaderSize.
  if (data.length < 8)
    throw new ParseError("Invalid RC4 FilePass record: missing encryption header")
  const flags = u32(data, 0)
  const headerSize = u32(data, 4)
  const headerStart = 8
  const verifierStart = headerStart + headerSize
  if (verifierStart + 36 > data.length) {
    throw new ParseError("Invalid RC4 CryptoAPI FilePass record: verifier extends past record")
  }

  const header = data.subarray(headerStart, Math.min(verifierStart, data.length))
  const algId = header.length >= 12 ? u32(header, 8) : 0x6801
  const algHash = header.length >= 16 ? u32(header, 12) : 0x8004
  const keyBits = header.length >= 20 ? u32(header, 16) : 40

  if (algId !== 0x6801) {
    throw new EncryptedFileError(
      "xls",
      `Unsupported XLS CryptoAPI cipher algId 0x${algId.toString(16)}; RC4 is supported.`,
    )
  }

  const hash = algHash === 0x8003 ? "md5" : "sha1"
  const saltSize = u32(data, verifierStart)
  if (saltSize <= 0 || saltSize > 64 || verifierStart + 4 + saltSize + 20 > data.length) {
    throw new ParseError("Invalid RC4 CryptoAPI FilePass verifier salt")
  }
  const salt = data.subarray(verifierStart + 4, verifierStart + 4 + saltSize)
  const encryptedVerifier = data.subarray(
    verifierStart + 4 + saltSize,
    verifierStart + 4 + saltSize + 16,
  )
  const hashSizeOffset = verifierStart + 4 + saltSize + 16
  const encryptedVerifierHashSize = u32(data, hashSizeOffset)
  const encryptedVerifierHash = data.subarray(hashSizeOffset + 4)

  return {
    method: (flags & 0x00000004) !== 0 ? "cryptoapi" : "rc4",
    salt,
    encryptedVerifier,
    encryptedVerifierHash,
    encryptedVerifierHashSize:
      encryptedVerifierHashSize > 0 ? encryptedVerifierHashSize : hashByteLength(hash),
    keySize: keyBits || 40,
    hash,
  }
}

function deriveRc4BlockKey(baseHash: Uint8Array, block: number, info: Rc4FilePassInfo): Uint8Array {
  if (info.method === "cryptoapi") {
    const h = digest(info.hash, concat([baseHash, int32le(block)]))
    return truncateRc4Key(h, info.keySize)
  }
  const h = md5(concat([baseHash.subarray(0, 5), int32le(block)]))
  return truncateRc4Key(h, info.keySize)
}

function truncateRc4Key(key: Uint8Array, keyBits: number): Uint8Array {
  const keyBytes = Math.max(1, Math.ceil(keyBits / 8))
  const out = new Uint8Array(keyBytes)
  out.set(key.subarray(0, Math.min(keyBytes, key.length)))
  return out
}

function hashByteLength(hash: "md5" | "sha1"): number {
  return hash === "md5" ? 16 : 20
}

function digest(hash: "md5" | "sha1", data: Uint8Array): Uint8Array {
  return hash === "md5" ? md5(data) : sha1(data)
}

function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  const s = rc4State(key)
  const out = new Uint8Array(data.length)
  for (let i = 0; i < data.length; i++) out[i] = data[i]! ^ rc4Next(s)
  return out
}

function rc4Drop(key: Uint8Array, drop: number, len: number): Uint8Array {
  const s = rc4State(key)
  for (let i = 0; i < drop; i++) rc4Next(s)
  const out = new Uint8Array(len)
  for (let i = 0; i < len; i++) out[i] = rc4Next(s)
  return out
}

function rc4State(key: Uint8Array): { s: Uint8Array; i: number; j: number } {
  const s = new Uint8Array(256)
  for (let i = 0; i < 256; i++) s[i] = i
  let j = 0
  for (let i = 0; i < 256; i++) {
    j = (j + s[i]! + key[i % key.length]!) & 0xff
    const t = s[i]!
    s[i] = s[j]!
    s[j] = t
  }
  return { s, i: 0, j: 0 }
}

function rc4Next(state: { s: Uint8Array; i: number; j: number }): number {
  state.i = (state.i + 1) & 0xff
  state.j = (state.j + state.s[state.i]!) & 0xff
  const t = state.s[state.i]!
  state.s[state.i] = state.s[state.j]!
  state.s[state.j] = t
  return state.s[(state.s[state.i]! + state.s[state.j]!) & 0xff]!
}

// ── MD5 ──────────────────────────────────────────────────────────────

function md5(input: Uint8Array): Uint8Array {
  const data = padHash(input, true)
  let a = 0x67452301
  let b = 0xefcdab89
  let c = 0x98badcfe
  let d = 0x10325476
  const s = [
    7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9,
    14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21,
    6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21,
  ]
  const k = new Uint32Array(64)
  for (let i = 0; i < 64; i++) k[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 0x100000000) >>> 0

  for (let off = 0; off < data.length; off += 64) {
    const m = new Uint32Array(16)
    for (let i = 0; i < 16; i++) m[i] = u32(data, off + i * 4)
    let aa = a,
      bb = b,
      cc = c,
      dd = d
    for (let i = 0; i < 64; i++) {
      let f: number
      let g: number
      if (i < 16) {
        f = (bb & cc) | (~bb & dd)
        g = i
      } else if (i < 32) {
        f = (dd & bb) | (~dd & cc)
        g = (5 * i + 1) % 16
      } else if (i < 48) {
        f = bb ^ cc ^ dd
        g = (3 * i + 5) % 16
      } else {
        f = cc ^ (bb | ~dd)
        g = (7 * i) % 16
      }
      const temp = dd
      dd = cc
      cc = bb
      bb = (bb + rotl32((aa + f + k[i]! + m[g]!) >>> 0, s[i]!)) >>> 0
      aa = temp
    }
    a = (a + aa) >>> 0
    b = (b + bb) >>> 0
    c = (c + cc) >>> 0
    d = (d + dd) >>> 0
  }
  return wordsLe([a, b, c, d])
}

// ── SHA-1 ────────────────────────────────────────────────────────────

function sha1(input: Uint8Array): Uint8Array {
  const data = padHash(input, false)
  let h0 = 0x67452301
  let h1 = 0xefcdab89
  let h2 = 0x98badcfe
  let h3 = 0x10325476
  let h4 = 0xc3d2e1f0
  const w = new Uint32Array(80)

  for (let off = 0; off < data.length; off += 64) {
    for (let i = 0; i < 16; i++) {
      w[i] =
        ((data[off + i * 4]! << 24) |
          (data[off + i * 4 + 1]! << 16) |
          (data[off + i * 4 + 2]! << 8) |
          data[off + i * 4 + 3]!) >>>
        0
    }
    for (let i = 16; i < 80; i++) w[i] = rotl32(w[i - 3]! ^ w[i - 8]! ^ w[i - 14]! ^ w[i - 16]!, 1)
    let a = h0,
      b = h1,
      c = h2,
      d = h3,
      e = h4
    for (let i = 0; i < 80; i++) {
      let f: number
      let k: number
      if (i < 20) {
        f = (b & c) | (~b & d)
        k = 0x5a827999
      } else if (i < 40) {
        f = b ^ c ^ d
        k = 0x6ed9eba1
      } else if (i < 60) {
        f = (b & c) | (b & d) | (c & d)
        k = 0x8f1bbcdc
      } else {
        f = b ^ c ^ d
        k = 0xca62c1d6
      }
      const temp = (rotl32(a, 5) + f + e + k + w[i]!) >>> 0
      e = d
      d = c
      c = rotl32(b, 30)
      b = a
      a = temp
    }
    h0 = (h0 + a) >>> 0
    h1 = (h1 + b) >>> 0
    h2 = (h2 + c) >>> 0
    h3 = (h3 + d) >>> 0
    h4 = (h4 + e) >>> 0
  }
  return wordsBe([h0, h1, h2, h3, h4])
}

function padHash(input: Uint8Array, littleEndianLength: boolean): Uint8Array {
  const bitLen = input.length * 8
  const total = ((input.length + 9 + 63) >> 6) << 6
  const out = new Uint8Array(total)
  out.set(input)
  out[input.length] = 0x80
  const dvOut = view(out)
  if (littleEndianLength) {
    dvOut.setUint32(total - 8, bitLen >>> 0, true)
    dvOut.setUint32(total - 4, Math.floor(bitLen / 0x100000000), true)
  } else {
    dvOut.setUint32(total - 8, Math.floor(bitLen / 0x100000000), false)
    dvOut.setUint32(total - 4, bitLen >>> 0, false)
  }
  return out
}

function rotl32(x: number, n: number): number {
  return ((x << n) | (x >>> (32 - n))) >>> 0
}

function wordsLe(words: number[]): Uint8Array {
  const out = new Uint8Array(words.length * 4)
  for (let i = 0; i < words.length; i++) view(out).setUint32(i * 4, words[i]!, true)
  return out
}

function wordsBe(words: number[]): Uint8Array {
  const out = new Uint8Array(words.length * 4)
  for (let i = 0; i < words.length; i++) view(out).setUint32(i * 4, words[i]!, false)
  return out
}
