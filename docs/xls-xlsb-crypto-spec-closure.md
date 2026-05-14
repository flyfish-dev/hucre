# XLS / XLSB spec-closure follow-up

This follow-up closes the most important gap from the previous implementation notes: encrypted OOXML packages used by `.xlsb` / `.xlsx` / `.xlsm` files.

## Spec areas rechecked

- **MS-XLS**: CFB-hosted BIFF workbook streams, workbook globals, worksheet substreams, BIFF8 strings, styles, formulas, hyperlinks, names, and document properties.
- **MS-XLSB**: OPC-hosted BIFF12 binary workbook / worksheet / shared-string / style records plus short cell records.
- **MS-CFB**: FAT/DIFAT directory and stream extraction; the reader now also tolerates simple regular-sector streams below the mini-stream cutoff when a producer omits MiniFAT.
- **MS-OFFCRYPTO**: Agile AES encrypted OOXML package envelopes with `EncryptionInfo` and `EncryptedPackage` streams.

## Added in this pass

- `src/crypto/office-crypto.ts`
  - detects Office encrypted CFB package envelopes
  - decrypts Agile AES `EncryptionInfo` + `EncryptedPackage` streams using `{ password }`
  - verifies passwords through encrypted verifier hash input/value
  - decrypts segmented 4096-byte encrypted package payloads
  - exports `encryptOfficeAgilePackageParts()` and `encryptOfficeAgilePackage()` for zero-dependency Agile AES encryption output
  - uses Web Crypto where available and Node `node:crypto` AES-CBC no-padding when available, avoiding npm dependencies
- `read()` now decrypts encrypted OOXML envelopes before dispatching to XLSX/XLSB/ODS/XLS readers.
- `readXlsb()` now decrypts encrypted XLSB envelopes directly when called with `{ password }`.
- `readXls()` now distinguishes real legacy XLS CFB files from encrypted OOXML package CFB files and reports decrypted-but-not-XLS content clearly.

## Coverage statement

This branch covers unencrypted XLS/BIFF and XLSB/BIFF12 workbook data parsing plus Agile AES encrypted OOXML package decryption/encryption with no npm dependency.

Legacy XLS `FilePass` encryption (XOR obfuscation / RC4 CryptoAPI applied inside BIFF record streams) is still detected and rejected with `EncryptedFileError`. I did **not** mark that path as complete because it needs fixture-backed record-level decryption and verification. Claiming that as complete without encrypted BIFF fixtures would be misleading.

## Validation performed in this bundle

- `git bundle verify` on the final artifact.
- `git clone <bundle> <dir>` checkout validation on the final artifact.
- TypeScript syntax/sanity check for the new crypto module and CFB reader.
- Agile encryption/decryption round-trip checks for payload lengths crossing the 4096-byte package segment boundary.
- Uploaded `.xls` fixture render artifacts from the previous pass remain included.
