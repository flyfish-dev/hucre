import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { streamXlsxRows } from "../src/xlsx/stream-reader"
import { read, readObjects } from "../src/defter"
import { isOle2Container } from "../src/_input"
import { DecryptionError, EncryptedFileError } from "../src/errors"
import { readCfb, writeCfb } from "../src/xlsx/crypto/cfb"
import { CfbReader } from "../src/xls/cfb"
import { decryptAgile, encryptAgile } from "../src/xlsx/crypto/agile"
import {
  decryptOfficeEncryptedPackage,
  encryptOfficeAgilePackage,
  encryptOfficeAgilePackageParts,
  isOfficeEncryptedPackage,
} from "../src/crypto/office-crypto"
import type { CellValue } from "../src/_types"

const FAST = { spinCount: 64 }

function book() {
  return {
    sheets: [
      {
        name: "Sheet1",
        rows: [
          ["Name", "Score"],
          ["Ada", 95],
          ["Linus", 88],
        ] as CellValue[][],
      },
    ],
  }
}

describe("CFB container", () => {
  it("round-trips mini (small) and regular (large) streams", () => {
    const small = new Uint8Array(1300).map((_, i) => i & 0xff)
    const big = new Uint8Array(20000).map((_, i) => (i * 31) & 0xff)
    const streams = readCfb(
      writeCfb([
        { name: "EncryptionInfo", data: small },
        { name: "EncryptedPackage", data: big },
      ]),
    )
    expect([...streams.get("EncryptionInfo")!]).toEqual([...small])
    expect([...streams.get("EncryptedPackage")!]).toEqual([...big])
  })
})

describe("agile crypto primitive", () => {
  it("encrypt → decrypt recovers arbitrary bytes", async () => {
    const payload = new TextEncoder().encode("PK" + "payload ".repeat(900))
    const enc = await encryptAgile(payload, "hunter2", FAST)
    expect(isOle2Container(enc)).toBe(true)
    expect([...(await decryptAgile(enc, "hunter2"))]).toEqual([...payload])
  })

  it("wrong password rejects with DecryptionError", async () => {
    const enc = await encryptAgile(new TextEncoder().encode("data ".repeat(2000)), "right", FAST)
    await expect(decryptAgile(enc, "wrong")).rejects.toBeInstanceOf(DecryptionError)
  })

  // 100,000 spins on each side is 200,000 SHA-512 iterations, and the
  // budget has to cover the worst case rather than the quiet one: vitest
  // runs files in parallel, so this shares a machine with the rest of the
  // suite. 30s was enough in isolation and not under load, so it failed
  // intermittently — on a test whose subject is not speed. Keeping the
  // real default matters: this is the only assertion that Excel's own
  // spin count works end to end.
  it("interoperates at Excel's default spin count", { timeout: 180_000 }, async () => {
    const payload = new TextEncoder().encode("z".repeat(9000))
    const enc = await encryptAgile(payload, "pw") // default 100000
    expect([...(await decryptAgile(enc, "pw"))]).toEqual([...payload])
  })
})

describe("shared Office Agile package", () => {
  const payload = Uint8Array.from({ length: 8_197 }, (_, index) => index % 251)

  for (const [keyBits, hashAlgorithm] of [
    [128, "SHA-1"],
    [192, "SHA-256"],
    [256, "SHA-384"],
    [256, "SHA-512"],
  ] as const) {
    it(`round-trips multiple encrypted segments with AES-${keyBits} and ${hashAlgorithm}`, async () => {
      const encrypted = await encryptOfficeAgilePackage(payload, {
        password: "secret",
        spinCount: 8,
        keyBits,
        hashAlgorithm,
      })
      expect(isOfficeEncryptedPackage(encrypted)).toBe(true)
      expect(new CfbReader(encrypted).getStream("EncryptedPackage")?.length).toBe(
        8 + 4_096 * 2 + 16,
      )
      expect(await decryptOfficeEncryptedPackage(encrypted, "secret", "xlsb")).toEqual(payload)
      await expect(
        decryptOfficeEncryptedPackage(encrypted, "wrong", "xlsb"),
      ).rejects.toBeInstanceOf(EncryptedFileError)
    })
  }

  it("keeps the two encrypted streams available to CFB package writers", async () => {
    const parts = await encryptOfficeAgilePackageParts(payload.subarray(0, 39), {
      password: "secret",
      spinCount: 8,
    })
    const container = writeCfb([
      { name: "EncryptionInfo", data: parts.encryptionInfo },
      { name: "EncryptedPackage", data: parts.encryptedPackage },
    ])
    expect(await decryptOfficeEncryptedPackage(container, "secret", "xls")).toEqual(
      payload.subarray(0, 39),
    )
    await expect(decryptOfficeEncryptedPackage(container, undefined, "xls")).rejects.toBeInstanceOf(
      EncryptedFileError,
    )
  })

  it("does not mistake arbitrary CFB content for an encrypted workbook", () => {
    expect(isOfficeEncryptedPackage(new Uint8Array(8))).toBe(false)
    expect(
      isOfficeEncryptedPackage(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])),
    ).toBe(false)
    expect(isOfficeEncryptedPackage(writeCfb([{ name: "Workbook", data: payload }]))).toBe(false)
  })
})

describe("writeXlsx encryption ↔ readXlsx decryption", () => {
  it("encrypts on write and decrypts on read with the password", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    expect(isOle2Container(enc)).toBe(true) // output is an encrypted OLE2 container, not a ZIP

    const wb = await readXlsx(enc, { password: "pw" })
    expect(wb.sheets[0].name).toBe("Sheet1")
    expect(wb.sheets[0].rows[1]).toEqual(["Ada", 95])
  })

  it("reading without a password throws EncryptedFileError", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    await expect(readXlsx(enc)).rejects.toBeInstanceOf(EncryptedFileError)
  })

  it("reading with the wrong password throws DecryptionError", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    await expect(readXlsx(enc, { password: "nope" })).rejects.toBeInstanceOf(DecryptionError)
  })
})

describe("decryption across the read entry points", () => {
  it("read() auto-detects and decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const wb = await read(enc, { password: "pw" })
    expect(wb.sheets[0].rows[2]).toEqual(["Linus", 88])
  })

  it("readObjects() decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const { data } = await readObjects<{ Name: string; Score: number }>(enc, { password: "pw" })
    expect(data).toEqual([
      { Name: "Ada", Score: 95 },
      { Name: "Linus", Score: 88 },
    ])
  })

  it("streamXlsxRows() decrypts", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const rows: CellValue[][] = []
    for await (const row of streamXlsxRows(enc, { password: "pw" })) rows.push(row.values)
    expect(rows[0]).toEqual(["Name", "Score"])
    expect(rows[1]).toEqual(["Ada", 95])
  })

  it("streamXlsxRows() without a password throws EncryptedFileError", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    await expect(async () => {
      for await (const _ of streamXlsxRows(enc)) void _
    }).rejects.toBeInstanceOf(EncryptedFileError)
  })
})

describe("roundtrip open → save with encryption", () => {
  it("opens an encrypted workbook and re-saves it encrypted", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "first", spinCount: 64 } })
    const wb = await openXlsx(enc, { password: "first" })
    expect(wb.sheets[0].rows[1]).toEqual(["Ada", 95])

    const resaved = await saveXlsx(wb, { encryption: { password: "second", spinCount: 64 } })
    expect(isOle2Container(resaved)).toBe(true)

    const reopened = await readXlsx(resaved, { password: "second" })
    expect(reopened.sheets[0].rows[2]).toEqual(["Linus", 88])
    // old password no longer works on the re-encrypted file
    await expect(readXlsx(resaved, { password: "first" })).rejects.toBeInstanceOf(DecryptionError)
  })

  it("saveXlsx without an encryption option produces a plain ZIP", async () => {
    const enc = await writeXlsx({ ...book(), encryption: { password: "pw", spinCount: 64 } })
    const wb = await openXlsx(enc, { password: "pw" })
    const plain = await saveXlsx(wb)
    expect(isOle2Container(plain)).toBe(false)
    expect((await readXlsx(plain)).sheets[0].rows[1]).toEqual(["Ada", 95])
  })
})
