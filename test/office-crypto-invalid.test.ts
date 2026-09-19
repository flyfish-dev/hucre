import { beforeAll, describe, expect, it } from "vitest"
import { DecryptionError, EncryptedFileError, ParseError } from "../src/errors"
import {
  decryptOfficeEncryptedPackage,
  encryptOfficeAgilePackageParts,
} from "../src/crypto/office-crypto"
import { writeCfb } from "../src/xlsx/crypto/cfb"

describe("Office Agile descriptor and payload validation", () => {
  const payload = Uint8Array.from({ length: 4_097 }, (_, index) => index % 239)
  let encryptionInfo: Uint8Array
  let encryptedPackage: Uint8Array

  beforeAll(async () => {
    const parts = await encryptOfficeAgilePackageParts(payload, {
      password: "pw",
      spinCount: 2,
    })
    encryptionInfo = parts.encryptionInfo
    encryptedPackage = parts.encryptedPackage
  })

  const packageWith = (info: Uint8Array, encrypted = encryptedPackage) =>
    writeCfb([
      { name: "EncryptionInfo", data: info },
      { name: "EncryptedPackage", data: encrypted },
    ])
  const changeXml = (change: (xml: string) => string) => {
    const xml = new TextDecoder().decode(encryptionInfo.subarray(8))
    const body = new TextEncoder().encode(change(xml))
    const info = new Uint8Array(8 + body.length)
    info.set(encryptionInfo.subarray(0, 8))
    info.set(body, 8)
    return info
  }

  it("requires both named encrypted streams and a complete descriptor header", async () => {
    const noPackage = writeCfb([{ name: "EncryptionInfo", data: encryptionInfo }])
    const noInfo = writeCfb([{ name: "EncryptedPackage", data: encryptedPackage }])
    await expect(decryptOfficeEncryptedPackage(noPackage, "pw")).rejects.toBeInstanceOf(ParseError)
    await expect(decryptOfficeEncryptedPackage(noInfo, "pw")).rejects.toBeInstanceOf(ParseError)
    await expect(
      decryptOfficeEncryptedPackage(packageWith(encryptionInfo.subarray(0, 7)), "pw"),
    ).rejects.toThrow(/too short/)
  })

  it("accepts XML under the extensible flag but rejects unsupported descriptor versions", async () => {
    const extensible = encryptionInfo.slice()
    new DataView(extensible.buffer).setUint16(0, 3, true)
    new DataView(extensible.buffer).setUint16(2, 2, true)
    new DataView(extensible.buffer).setUint32(4, 0x10, true)
    expect(await decryptOfficeEncryptedPackage(packageWith(extensible), "pw")).toEqual(payload)
    new DataView(extensible.buffer).setUint32(4, 0, true)
    await expect(
      decryptOfficeEncryptedPackage(packageWith(extensible), "pw"),
    ).rejects.toBeInstanceOf(EncryptedFileError)
  })

  it("rejects missing keys, unsafe spin counts and unsupported hashes before deriving a key", async () => {
    const variants = [
      [changeXml((xml) => xml.replace(/<keyData\b[^>]*\/>/, "")), ParseError],
      [changeXml((xml) => xml.replace(/spinCount="\d+"/, 'spinCount="-1"')), DecryptionError],
      [
        changeXml((xml) => xml.replace(/spinCount="\d+"/, 'spinCount="1000000000"')),
        DecryptionError,
      ],
      [
        changeXml((xml) => xml.replace(/hashAlgorithm="[^"]+"/, 'hashAlgorithm="SHA999"')),
        ParseError,
      ],
      [changeXml((xml) => xml.replace(/ saltValue="[^"]+"/, "")), ParseError],
    ] as const
    for (const [info, error] of variants) {
      await expect(decryptOfficeEncryptedPackage(packageWith(info), "pw")).rejects.toBeInstanceOf(
        error,
      )
    }
  })

  it("rejects a short or truncated encrypted package after password verification", async () => {
    await expect(
      decryptOfficeEncryptedPackage(
        packageWith(encryptionInfo, encryptedPackage.subarray(0, 7)),
        "pw",
      ),
    ).rejects.toThrow(/too short/)
    await expect(
      decryptOfficeEncryptedPackage(
        packageWith(encryptionInfo, encryptedPackage.subarray(0, -1)),
        "pw",
      ),
    ).rejects.toThrow(/truncated/)
  })
})
