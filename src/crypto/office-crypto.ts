// ── Office Crypto Helpers ────────────────────────────────────────────
// Minimal-dependency MS-OFFCRYPTO implementation for encrypted OOXML
// packages (.xlsx/.xlsb/.xlsm stored as an OLE2/CFB envelope with
// EncryptionInfo + EncryptedPackage streams). The implementation uses only
// Web Crypto primitives available in modern browsers and Node 20+.

import { DecryptionError, EncryptedFileError, ParseError } from "../errors"
import type { WorkbookFormat } from "../errors"
import { MAX_SPIN_COUNT } from "../limits"
import { CfbReader } from "../xls/cfb"

export interface OfficeCryptoOptions {
  password?: string
}

export interface EncryptedOfficePackageParts {
  encryptionInfo: Uint8Array
  encryptedPackage: Uint8Array
}

export interface AgileEncryptionOptions {
  password: string
  spinCount?: number
  keyBits?: 128 | 192 | 256
  hashAlgorithm?: "SHA-1" | "SHA-256" | "SHA-384" | "SHA-512"
}

const ENCRYPTED_PACKAGE_STREAM = "EncryptedPackage"
const ENCRYPTION_INFO_STREAM = "EncryptionInfo"
const PACKAGE_SEGMENT_SIZE = 4096

const BLOCKKEY_VERIFIER_INPUT = bytes([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79])
const BLOCKKEY_VERIFIER_VALUE = bytes([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e])
const BLOCKKEY_ENCRYPTED_KEY = bytes([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6])

const CFB_MAGIC = Object.freeze([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const)
const END_OF_CHAIN = 0xfffffffe
const FAT_SECT = 0xfffffffd
const FREE_SECT = 0xffffffff

function bytes(values: number[]): Uint8Array {
  return new Uint8Array(values)
}

function isCfb(data: Uint8Array): boolean {
  if (data.length < CFB_MAGIC.length) return false
  for (let i = 0; i < CFB_MAGIC.length; i++) if (data[i] !== CFB_MAGIC[i]) return false
  return true
}

export function isOfficeEncryptedPackage(data: Uint8Array): boolean {
  if (!isCfb(data)) return false
  try {
    const cfb = new CfbReader(data)
    return cfb.hasStream(ENCRYPTION_INFO_STREAM) && cfb.hasStream(ENCRYPTED_PACKAGE_STREAM)
  } catch {
    return false
  }
}

export async function decryptOfficeEncryptedPackage(
  data: Uint8Array,
  password: string | undefined,
  format?: WorkbookFormat,
): Promise<Uint8Array> {
  if (!password) {
    throw new EncryptedFileError(
      format,
      'File is password-protected. Pass `{ password: "..." }` in read options to decrypt it.',
    )
  }

  const cfb = new CfbReader(data)
  const encryptionInfo = cfb.getStream(ENCRYPTION_INFO_STREAM)
  const encryptedPackage = cfb.getStream(ENCRYPTED_PACKAGE_STREAM)
  if (!encryptionInfo || !encryptedPackage) {
    throw new ParseError(
      "Invalid encrypted Office package: missing EncryptionInfo or EncryptedPackage stream",
    )
  }

  return decryptEncryptionInfoPackage(encryptionInfo, encryptedPackage, password, format)
}

async function decryptEncryptionInfoPackage(
  encryptionInfo: Uint8Array,
  encryptedPackage: Uint8Array,
  password: string,
  format?: WorkbookFormat,
): Promise<Uint8Array> {
  if (encryptionInfo.length < 8) throw new ParseError("Invalid EncryptionInfo stream: too short")
  const view = dv(encryptionInfo)
  const major = view.getUint16(0, true)
  const minor = view.getUint16(2, true)
  const flags = view.getUint32(4, true)

  // Agile EncryptionInfo is version 4.4 and stores XML after the 8-byte header.
  if (major === 4 && minor === 4) {
    return decryptAgilePackage(encryptionInfo.subarray(8), encryptedPackage, password, format)
  }

  // Extensible encryption also uses an XML descriptor. A few producers set
  // the extensible flag but still emit the same XML grammar; support that
  // shape rather than forcing callers through an error.
  if ((flags & 0x10) !== 0) {
    return decryptAgilePackage(encryptionInfo.subarray(8), encryptedPackage, password, format)
  }

  throw new EncryptedFileError(
    format,
    `Unsupported Office encryption descriptor version ${major}.${minor}. Agile AES encryption is supported; legacy Standard/CryptoAPI BIFF FilePass remains unsupported without a fixture-backed implementation.`,
  )
}

interface AgileKeyData {
  saltValue: Uint8Array
  blockSize: number
  keyBits: number
  hashSize: number
  cipherAlgorithm: string
  cipherChaining: string
  hashAlgorithm: HashAlgorithmIdentifier
}

interface AgileEncryptedKey {
  saltValue: Uint8Array
  spinCount: number
  blockSize: number
  keyBits: number
  hashSize: number
  cipherAlgorithm: string
  cipherChaining: string
  hashAlgorithm: HashAlgorithmIdentifier
  encryptedVerifierHashInput: Uint8Array
  encryptedVerifierHashValue: Uint8Array
  encryptedKeyValue: Uint8Array
}

interface AgileInfo {
  keyData: AgileKeyData
  encryptedKey: AgileEncryptedKey
}

async function decryptAgilePackage(
  xmlBytes: Uint8Array,
  encryptedPackage: Uint8Array,
  password: string,
  format?: WorkbookFormat,
): Promise<Uint8Array> {
  try {
    return await decryptAgilePackageUnchecked(xmlBytes, encryptedPackage, password, format)
  } catch (err) {
    if (
      err instanceof DecryptionError ||
      err instanceof EncryptedFileError ||
      err instanceof ParseError
    )
      throw err
    throw new EncryptedFileError(
      format,
      "Incorrect password or corrupted encrypted Office workbook.",
    )
  }
}

async function decryptAgilePackageUnchecked(
  xmlBytes: Uint8Array,
  encryptedPackage: Uint8Array,
  password: string,
  format?: WorkbookFormat,
): Promise<Uint8Array> {
  const xml = decodeUtf8(stripTrailingZeros(xmlBytes))
  const info = parseAgileInfo(xml)
  const passwordHash = await hashPasswordAgile(
    password,
    info.encryptedKey.saltValue,
    info.encryptedKey.spinCount,
    info.encryptedKey.hashAlgorithm,
  )

  const verifierKey = await deriveAgileKey(
    passwordHash,
    BLOCKKEY_VERIFIER_INPUT,
    info.encryptedKey.hashAlgorithm,
    info.encryptedKey.keyBits / 8,
  )
  const verifierIv = await deriveAgileIv(
    info.encryptedKey.saltValue,
    BLOCKKEY_VERIFIER_INPUT,
    info.encryptedKey.hashAlgorithm,
    info.encryptedKey.blockSize,
  )
  const verifierInput = await aesCbcDecrypt(
    verifierKey,
    verifierIv,
    info.encryptedKey.encryptedVerifierHashInput,
  )

  const verifierValueKey = await deriveAgileKey(
    passwordHash,
    BLOCKKEY_VERIFIER_VALUE,
    info.encryptedKey.hashAlgorithm,
    info.encryptedKey.keyBits / 8,
  )
  const verifierValueIv = await deriveAgileIv(
    info.encryptedKey.saltValue,
    BLOCKKEY_VERIFIER_VALUE,
    info.encryptedKey.hashAlgorithm,
    info.encryptedKey.blockSize,
  )
  const verifierHashValue = await aesCbcDecrypt(
    verifierValueKey,
    verifierValueIv,
    info.encryptedKey.encryptedVerifierHashValue,
  )
  const verifierHash = await digest(info.encryptedKey.hashAlgorithm, verifierInput)
  if (
    !constantTimeStartsWith(verifierHashValue, verifierHash.subarray(0, info.encryptedKey.hashSize))
  ) {
    throw new EncryptedFileError(format, "Incorrect password for encrypted Office workbook.")
  }

  const secretKeyKey = await deriveAgileKey(
    passwordHash,
    BLOCKKEY_ENCRYPTED_KEY,
    info.encryptedKey.hashAlgorithm,
    info.encryptedKey.keyBits / 8,
  )
  const secretKeyIv = await deriveAgileIv(
    info.encryptedKey.saltValue,
    BLOCKKEY_ENCRYPTED_KEY,
    info.encryptedKey.hashAlgorithm,
    info.encryptedKey.blockSize,
  )
  const secretKey = (
    await aesCbcDecrypt(secretKeyKey, secretKeyIv, info.encryptedKey.encryptedKeyValue)
  ).subarray(0, info.keyData.keyBits / 8)

  if (encryptedPackage.length < 8)
    throw new ParseError("Invalid EncryptedPackage stream: too short")
  const packageView = dv(encryptedPackage)
  const originalSize = Number(packageView.getBigUint64(0, true))
  const out = new Uint8Array(originalSize)
  let encryptedOffset = 8
  let plainOffset = 0
  let block = 0

  while (plainOffset < originalSize) {
    const plainLen = Math.min(PACKAGE_SEGMENT_SIZE, originalSize - plainOffset)
    const encryptedLen = align(plainLen, info.keyData.blockSize)
    const encryptedChunk = encryptedPackage.subarray(
      encryptedOffset,
      encryptedOffset + encryptedLen,
    )
    if (encryptedChunk.length !== encryptedLen) {
      throw new ParseError("Invalid EncryptedPackage stream: encrypted payload is truncated")
    }
    const iv = await deriveAgileIv(
      info.keyData.saltValue,
      int32le(block),
      info.keyData.hashAlgorithm,
      info.keyData.blockSize,
    )
    const plainChunk = await aesCbcDecrypt(secretKey, iv, encryptedChunk)
    out.set(plainChunk.subarray(0, plainLen), plainOffset)
    encryptedOffset += encryptedLen
    plainOffset += plainLen
    block++
  }

  return out
}

export async function encryptOfficeAgilePackageParts(
  packageData: Uint8Array,
  options: AgileEncryptionOptions,
): Promise<EncryptedOfficePackageParts> {
  const keyBits = options.keyBits ?? 256
  const hashAlgorithm = options.hashAlgorithm ?? "SHA-512"
  const spinCount = options.spinCount ?? 100000
  const blockSize = 16
  const hashSize = hashByteLength(hashAlgorithm)
  const keyDataSalt = randomBytes(16)
  const encryptedKeySalt = randomBytes(16)
  const packageKey = randomBytes(keyBits / 8)
  const verifierInputPlain = randomBytes(16)
  const passwordHash = await hashPasswordAgile(
    options.password,
    encryptedKeySalt,
    spinCount,
    hashAlgorithm,
  )

  const verifierHash = await digest(hashAlgorithm, verifierInputPlain)
  const verifierInputKey = await deriveAgileKey(
    passwordHash,
    BLOCKKEY_VERIFIER_INPUT,
    hashAlgorithm,
    keyBits / 8,
  )
  const verifierInputIv = await deriveAgileIv(
    encryptedKeySalt,
    BLOCKKEY_VERIFIER_INPUT,
    hashAlgorithm,
    blockSize,
  )
  const encryptedVerifierHashInput = await aesCbcEncrypt(
    verifierInputKey,
    verifierInputIv,
    verifierInputPlain,
  )

  const verifierValueKey = await deriveAgileKey(
    passwordHash,
    BLOCKKEY_VERIFIER_VALUE,
    hashAlgorithm,
    keyBits / 8,
  )
  const verifierValueIv = await deriveAgileIv(
    encryptedKeySalt,
    BLOCKKEY_VERIFIER_VALUE,
    hashAlgorithm,
    blockSize,
  )
  const encryptedVerifierHashValue = await aesCbcEncrypt(
    verifierValueKey,
    verifierValueIv,
    verifierHash,
  )

  const secretKeyKey = await deriveAgileKey(
    passwordHash,
    BLOCKKEY_ENCRYPTED_KEY,
    hashAlgorithm,
    keyBits / 8,
  )
  const secretKeyIv = await deriveAgileIv(
    encryptedKeySalt,
    BLOCKKEY_ENCRYPTED_KEY,
    hashAlgorithm,
    blockSize,
  )
  const encryptedKeyValue = await aesCbcEncrypt(secretKeyKey, secretKeyIv, packageKey)

  const encryptedPayloadParts: Uint8Array[] = []
  for (
    let offset = 0, block = 0;
    offset < packageData.length;
    offset += PACKAGE_SEGMENT_SIZE, block++
  ) {
    const plainChunk = packageData.subarray(
      offset,
      Math.min(offset + PACKAGE_SEGMENT_SIZE, packageData.length),
    )
    const iv = await deriveAgileIv(keyDataSalt, int32le(block), hashAlgorithm, blockSize)
    encryptedPayloadParts.push(await aesCbcEncrypt(packageKey, iv, plainChunk))
  }

  const encryptedPackage = concat([uint64le(packageData.length), ...encryptedPayloadParts])
  const xml = buildAgileXml({
    keyDataSalt,
    encryptedKeySalt,
    spinCount,
    keyBits,
    blockSize,
    hashSize,
    hashAlgorithm,
    encryptedVerifierHashInput,
    encryptedVerifierHashValue,
    encryptedKeyValue,
  })
  const xmlBytes = encodeUtf8(xml)
  const encryptionInfo = concat([u16le(4), u16le(4), u32le(0x40), xmlBytes])
  return { encryptionInfo, encryptedPackage }
}

export async function encryptOfficeAgilePackage(
  packageData: Uint8Array,
  options: AgileEncryptionOptions,
): Promise<Uint8Array> {
  const parts = await encryptOfficeAgilePackageParts(packageData, options)
  return writeSimpleCfb([
    { name: ENCRYPTION_INFO_STREAM, data: parts.encryptionInfo },
    { name: ENCRYPTED_PACKAGE_STREAM, data: parts.encryptedPackage },
  ])
}

function parseAgileInfo(xml: string): AgileInfo {
  const keyDataTag = tagAttrs(xml, "keyData")
  const encryptedKeyTag = tagAttrs(xml, "encryptedKey")
  if (!keyDataTag || !encryptedKeyTag) {
    throw new ParseError(
      "Invalid Agile EncryptionInfo XML: missing keyData or encryptedKey element",
    )
  }

  const spinCount = intAttr(encryptedKeyTag, "spinCount", 100000)
  if (!Number.isSafeInteger(spinCount) || spinCount < 0) {
    throw new DecryptionError("EncryptionInfo has an invalid spinCount.")
  }
  if (spinCount > MAX_SPIN_COUNT) {
    throw new DecryptionError(
      `EncryptionInfo spinCount ${spinCount} exceeds the maximum of ${MAX_SPIN_COUNT}.`,
    )
  }

  return {
    keyData: {
      saltValue: b64(required(keyDataTag, "saltValue", "keyData")),
      blockSize: intAttr(keyDataTag, "blockSize", 16),
      keyBits: intAttr(keyDataTag, "keyBits", 256),
      hashSize: intAttr(
        keyDataTag,
        "hashSize",
        hashByteLength(normalizeHash(required(keyDataTag, "hashAlgorithm", "keyData"))),
      ),
      cipherAlgorithm: required(keyDataTag, "cipherAlgorithm", "keyData"),
      cipherChaining: required(keyDataTag, "cipherChaining", "keyData"),
      hashAlgorithm: normalizeHash(required(keyDataTag, "hashAlgorithm", "keyData")),
    },
    encryptedKey: {
      saltValue: b64(required(encryptedKeyTag, "saltValue", "encryptedKey")),
      spinCount,
      blockSize: intAttr(encryptedKeyTag, "blockSize", 16),
      keyBits: intAttr(encryptedKeyTag, "keyBits", 256),
      hashSize: intAttr(
        encryptedKeyTag,
        "hashSize",
        hashByteLength(normalizeHash(required(encryptedKeyTag, "hashAlgorithm", "encryptedKey"))),
      ),
      cipherAlgorithm: required(encryptedKeyTag, "cipherAlgorithm", "encryptedKey"),
      cipherChaining: required(encryptedKeyTag, "cipherChaining", "encryptedKey"),
      hashAlgorithm: normalizeHash(required(encryptedKeyTag, "hashAlgorithm", "encryptedKey")),
      encryptedVerifierHashInput: b64(
        required(encryptedKeyTag, "encryptedVerifierHashInput", "encryptedKey"),
      ),
      encryptedVerifierHashValue: b64(
        required(encryptedKeyTag, "encryptedVerifierHashValue", "encryptedKey"),
      ),
      encryptedKeyValue: b64(required(encryptedKeyTag, "encryptedKeyValue", "encryptedKey")),
    },
  }
}

function tagAttrs(xml: string, localName: string): Record<string, string> | null {
  const re = new RegExp(`<[^>:/]*:?${localName}\\b([^>]*)>`, "i")
  const m = xml.match(re)
  if (!m) return null
  const attrs: Record<string, string> = {}
  const attrText = m[1] ?? ""
  attrText.replace(/([:\w-]+)\s*=\s*"([^"]*)"/g, (_all, key: string, value: string) => {
    attrs[key.includes(":") ? key.slice(key.indexOf(":") + 1) : key] = value
    return ""
  })
  return attrs
}

function required(attrs: Record<string, string>, name: string, tag: string): string {
  const value = attrs[name]
  if (value === undefined)
    throw new ParseError(`Invalid Agile EncryptionInfo XML: ${tag}/@${name} is missing`)
  return value
}

function intAttr(attrs: Record<string, string>, name: string, fallback: number): number {
  const raw = attrs[name]
  if (raw === undefined || raw === "") return fallback
  const n = Number(raw)
  return Number.isFinite(n) ? n : fallback
}

function normalizeHash(value: string): HashAlgorithmIdentifier {
  const v = value.replace(/_/g, "-").toUpperCase()
  switch (v) {
    case "SHA1":
    case "SHA-1":
      return "SHA-1"
    case "SHA256":
    case "SHA-256":
      return "SHA-256"
    case "SHA384":
    case "SHA-384":
      return "SHA-384"
    case "SHA512":
    case "SHA-512":
      return "SHA-512"
    default:
      throw new ParseError(`Unsupported Agile hash algorithm: ${value}`)
  }
}

function hashByteLength(algorithm: HashAlgorithmIdentifier): number {
  switch (algorithm) {
    case "SHA-1":
      return 20
    case "SHA-256":
      return 32
    case "SHA-384":
      return 48
    case "SHA-512":
      return 64
    default:
      return 64
  }
}

async function hashPasswordAgile(
  password: string,
  salt: Uint8Array,
  spinCount: number,
  algorithm: HashAlgorithmIdentifier,
): Promise<Uint8Array> {
  let h = await digest(algorithm, concat([salt, utf16le(password)]))
  for (let i = 0; i < spinCount; i++) {
    h = await digest(algorithm, concat([int32le(i), h]))
  }
  return h
}

async function deriveAgileKey(
  passwordHash: Uint8Array,
  blockKey: Uint8Array,
  algorithm: HashAlgorithmIdentifier,
  keyBytes: number,
): Promise<Uint8Array> {
  const h = await digest(algorithm, concat([passwordHash, blockKey]))
  return padOrTruncate(h, keyBytes)
}

async function deriveAgileIv(
  salt: Uint8Array,
  blockKey: Uint8Array,
  algorithm: HashAlgorithmIdentifier,
  blockSize: number,
): Promise<Uint8Array> {
  if (blockKey.length === 0) return padOrTruncate(salt, blockSize)
  const h = await digest(algorithm, concat([salt, blockKey]))
  return padOrTruncate(h, blockSize)
}

function padOrTruncate(data: Uint8Array, len: number): Uint8Array {
  if (data.length === len) return data
  const out = new Uint8Array(len)
  out.fill(0x36)
  out.set(data.subarray(0, Math.min(data.length, len)))
  return out
}

async function digest(algorithm: HashAlgorithmIdentifier, data: Uint8Array): Promise<Uint8Array> {
  const subtle = cryptoSubtle()
  const out = await subtle.digest(algorithm, data as unknown as BufferSource)
  return new Uint8Array(out)
}

async function aesCbcDecrypt(
  keyBytes: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  const node = await nodeCrypto()
  if (node) {
    const buffer = nodeBuffer()
    const decipher = node.createDecipheriv(
      `aes-${keyBytes.length * 8}-cbc`,
      buffer.from(keyBytes),
      buffer.from(iv),
    )
    decipher.setAutoPadding(false)
    return new Uint8Array(buffer.concat([decipher.update(buffer.from(data)), decipher.final()]))
  }

  const subtle = cryptoSubtle()
  const key = await subtle.importKey(
    "raw",
    keyBytes as unknown as BufferSource,
    { name: "AES-CBC" },
    false,
    ["decrypt"],
  )
  const out = await subtle.decrypt(
    { name: "AES-CBC", iv: iv as unknown as BufferSource },
    key,
    data as unknown as BufferSource,
  )
  return new Uint8Array(out)
}

async function aesCbcEncrypt(
  keyBytes: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  const node = await nodeCrypto()
  if (node) {
    const buffer = nodeBuffer()
    const cipher = node.createCipheriv(
      `aes-${keyBytes.length * 8}-cbc`,
      buffer.from(keyBytes),
      buffer.from(iv),
    )
    cipher.setAutoPadding(false)
    const padded = data.length % 16 === 0 ? data : padOrTruncate(data, align(data.length, 16))
    return new Uint8Array(buffer.concat([cipher.update(buffer.from(padded)), cipher.final()]))
  }

  const subtle = cryptoSubtle()
  const key = await subtle.importKey(
    "raw",
    keyBytes as unknown as BufferSource,
    { name: "AES-CBC" },
    false,
    ["encrypt"],
  )
  const out = await subtle.encrypt(
    { name: "AES-CBC", iv: iv as unknown as BufferSource },
    key,
    data as unknown as BufferSource,
  )
  return new Uint8Array(out)
}

async function nodeCrypto(): Promise<null | {
  createCipheriv(
    algorithm: string,
    key: unknown,
    iv: unknown,
  ): { setAutoPadding(value: boolean): void; update(data: unknown): unknown; final(): unknown }
  createDecipheriv(
    algorithm: string,
    key: unknown,
    iv: unknown,
  ): { setAutoPadding(value: boolean): void; update(data: unknown): unknown; final(): unknown }
}> {
  const g = globalThis as unknown as { process?: { versions?: { node?: string } } }
  if (!g.process?.versions?.node) return null
  try {
    const importer = Function("specifier", "return import(specifier)") as (
      specifier: string,
    ) => Promise<unknown>
    return (await importer("node:crypto")) as Awaited<ReturnType<typeof nodeCrypto>>
  } catch {
    return null
  }
}

function nodeBuffer(): { from(data: Uint8Array): unknown; concat(parts: unknown[]): Uint8Array } {
  const ctor = (
    globalThis as unknown as {
      Buffer?: { from(data: Uint8Array): unknown; concat(parts: unknown[]): Uint8Array }
    }
  ).Buffer
  if (!ctor) throw new ParseError("Node Buffer is not available for AES-CBC no-padding.")
  return ctor
}

function cryptoSubtle(): SubtleCrypto {
  const subtle = globalThis.crypto?.subtle
  if (!subtle) {
    throw new ParseError(
      "Office decryption requires Web Crypto SubtleCrypto (available in browsers and Node 20+).",
    )
  }
  return subtle
}

function randomBytes(len: number): Uint8Array {
  if (!globalThis.crypto?.getRandomValues)
    throw new ParseError("Office encryption requires crypto.getRandomValues().")
  const out = new Uint8Array(len)
  globalThis.crypto.getRandomValues(out)
  return out
}

function constantTimeStartsWith(actual: Uint8Array, expected: Uint8Array): boolean {
  let diff = actual.length < expected.length ? 1 : 0
  for (let i = 0; i < expected.length; i++) diff |= (actual[i] ?? 0) ^ (expected[i] ?? 0)
  return diff === 0
}

function stripTrailingZeros(data: Uint8Array): Uint8Array {
  let end = data.length
  while (end > 0 && data[end - 1] === 0) end--
  return data.subarray(0, end)
}

function align(n: number, block: number): number {
  return Math.ceil(n / block) * block
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

function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2)
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    out[i * 2] = c & 0xff
    out[i * 2 + 1] = c >>> 8
  }
  return out
}

function encodeUtf8(s: string): Uint8Array {
  return new TextEncoder().encode(s)
}

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data)
}

function b64(s: string): Uint8Array {
  const atobFn = globalThis.atob
  if (typeof atobFn === "function") {
    const raw = atobFn(s)
    const out = new Uint8Array(raw.length)
    for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i)
    return out
  }
  const bufferCtor = (
    globalThis as unknown as { Buffer?: { from(data: string, encoding: string): Uint8Array } }
  ).Buffer
  if (bufferCtor) return new Uint8Array(bufferCtor.from(s, "base64"))
  throw new ParseError("Base64 decoding is not available in this runtime.")
}

function toB64(data: Uint8Array): string {
  const btoaFn = globalThis.btoa
  if (typeof btoaFn === "function") {
    let s = ""
    for (const b of data) s += String.fromCharCode(b)
    return btoaFn(s)
  }
  const bufferCtor = (
    globalThis as unknown as {
      Buffer?: { from(data: Uint8Array): { toString(enc: string): string } }
    }
  ).Buffer
  if (bufferCtor) return bufferCtor.from(data).toString("base64")
  throw new ParseError("Base64 encoding is not available in this runtime.")
}

function int32le(n: number): Uint8Array {
  const out = new Uint8Array(4)
  dv(out).setUint32(0, n, true)
  return out
}

function uint64le(n: number): Uint8Array {
  const out = new Uint8Array(8)
  dv(out).setBigUint64(0, BigInt(n), true)
  return out
}

function u16le(n: number): Uint8Array {
  const out = new Uint8Array(2)
  dv(out).setUint16(0, n, true)
  return out
}

function u32le(n: number): Uint8Array {
  const out = new Uint8Array(4)
  dv(out).setUint32(0, n, true)
  return out
}

function dv(data: Uint8Array): DataView {
  return new DataView(data.buffer, data.byteOffset, data.byteLength)
}

function buildAgileXml(args: {
  keyDataSalt: Uint8Array
  encryptedKeySalt: Uint8Array
  spinCount: number
  keyBits: number
  blockSize: number
  hashSize: number
  hashAlgorithm: HashAlgorithmIdentifier
  encryptedVerifierHashInput: Uint8Array
  encryptedVerifierHashValue: Uint8Array
  encryptedKeyValue: Uint8Array
}): string {
  const hash = String(args.hashAlgorithm).replace("-", "")
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<keyData saltSize="${args.keyDataSalt.length}" blockSize="${args.blockSize}" keyBits="${args.keyBits}" hashSize="${args.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${hash}" saltValue="${toB64(args.keyDataSalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="" encryptedHmacValue=""/>` +
    `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<p:encryptedKey spinCount="${args.spinCount}" saltSize="${args.encryptedKeySalt.length}" blockSize="${args.blockSize}" keyBits="${args.keyBits}" hashSize="${args.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${hash}" saltValue="${toB64(args.encryptedKeySalt)}" encryptedVerifierHashInput="${toB64(args.encryptedVerifierHashInput)}" encryptedVerifierHashValue="${toB64(args.encryptedVerifierHashValue)}" encryptedKeyValue="${toB64(args.encryptedKeyValue)}"/>` +
    `</keyEncryptor></keyEncryptors></encryption>`
  )
}

interface CfbWriteStream {
  name: string
  data: Uint8Array
}

function writeSimpleCfb(streams: CfbWriteStream[]): Uint8Array {
  const sectorSize = 512
  const directoryEntries = 1 + streams.length
  const directorySectors = Math.ceil((directoryEntries * 128) / sectorSize)
  const streamSectorCounts = streams.map((s) => Math.max(1, Math.ceil(s.data.length / sectorSize)))
  let nonFatSectors = directorySectors + streamSectorCounts.reduce((sum, n) => sum + n, 0)
  let fatSectors = 1
  for (;;) {
    const entriesNeeded = nonFatSectors + fatSectors
    const nextFatSectors = Math.ceil(entriesNeeded / (sectorSize / 4))
    if (nextFatSectors === fatSectors) break
    fatSectors = nextFatSectors
  }

  const totalSectors = nonFatSectors + fatSectors
  const out = new Uint8Array(512 + totalSectors * sectorSize)
  const view = dv(out)
  out.set(CFB_MAGIC, 0)
  view.setUint16(0x18, 0x003e, true)
  view.setUint16(0x1a, 0x0003, true)
  view.setUint16(0x1c, 0xfffe, true)
  view.setUint16(0x1e, 9, true)
  view.setUint16(0x20, 6, true)
  view.setUint32(0x2c, fatSectors, true)
  const firstDirectorySector = 0
  view.setUint32(0x30, firstDirectorySector, true)
  view.setUint32(0x38, 0, true) // no MiniFAT; small streams are regular streams
  view.setUint32(0x3c, END_OF_CHAIN, true)
  view.setUint32(0x40, 0, true)
  view.setUint32(0x44, END_OF_CHAIN, true)
  view.setUint32(0x48, 0, true)
  for (let i = 0; i < 109; i++) view.setUint32(0x4c + i * 4, FREE_SECT, true)

  const fat: number[] = new Array(totalSectors).fill(FREE_SECT)
  const directoryStart = 0
  for (let i = 0; i < directorySectors; i++) {
    fat[directoryStart + i] = i === directorySectors - 1 ? END_OF_CHAIN : directoryStart + i + 1
  }
  let sectorCursor = directorySectors
  const streamStarts: number[] = []
  for (let i = 0; i < streams.length; i++) {
    const start = sectorCursor
    streamStarts.push(start)
    const count = streamSectorCounts[i]!
    for (let j = 0; j < count; j++) fat[start + j] = j === count - 1 ? END_OF_CHAIN : start + j + 1
    out.set(streams[i]!.data, 512 + start * sectorSize)
    sectorCursor += count
  }
  const firstFatSector = sectorCursor
  for (let i = 0; i < fatSectors; i++) {
    fat[firstFatSector + i] = FAT_SECT
    view.setUint32(0x4c + i * 4, firstFatSector + i, true)
  }

  writeDirectoryEntry(
    out,
    directoryStart,
    0,
    "Root Entry",
    5,
    END_OF_CHAIN,
    0,
    streams.length > 0 ? 1 : 0,
    FREE_SECT,
    FREE_SECT,
  )
  for (let i = 0; i < streams.length; i++) {
    writeDirectoryEntry(
      out,
      directoryStart,
      i + 1,
      streams[i]!.name,
      2,
      streamStarts[i]!,
      streams[i]!.data.length,
      FREE_SECT,
      i === 0 ? FREE_SECT : i,
      i === streams.length - 1 ? FREE_SECT : i + 2,
    )
  }

  let fatEntryOffset = 512 + firstFatSector * sectorSize
  for (let i = 0; i < fat.length; i++) {
    view.setUint32(fatEntryOffset, fat[i]!, true)
    fatEntryOffset += 4
  }
  while (fatEntryOffset < 512 + (firstFatSector + fatSectors) * sectorSize) {
    view.setUint32(fatEntryOffset, FREE_SECT, true)
    fatEntryOffset += 4
  }

  return out
}

function writeDirectoryEntry(
  out: Uint8Array,
  directoryStartSector: number,
  index: number,
  name: string,
  type: number,
  startSector: number,
  size: number,
  child: number,
  left: number,
  right: number,
): void {
  const sectorSize = 512
  const pos = 512 + directoryStartSector * sectorSize + index * 128
  const view = dv(out)
  const nameBytes = utf16le(name + "\0")
  out.set(nameBytes.subarray(0, 64), pos)
  view.setUint16(pos + 64, Math.min(nameBytes.length, 64), true)
  out[pos + 66] = type
  out[pos + 67] = 1 // black
  view.setUint32(pos + 68, left, true)
  view.setUint32(pos + 72, right, true)
  view.setUint32(pos + 76, child, true)
  view.setUint32(pos + 116, startSector, true)
  view.setUint32(pos + 120, size >>> 0, true)
  view.setUint32(pos + 124, Math.floor(size / 0x1_0000_0000), true)
}
