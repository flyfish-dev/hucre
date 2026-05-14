# XLS / XLSB remaining-gap closure and validation

This pass rechecks every item in the uploaded `hucre XLS / XLSB read support notes` remaining-gap list and records the implementation / validation status.

## Remaining-gap checklist

| Uploaded checklist item | Status in this bundle | Evidence / implementation |
| --- | --- | --- |
| encrypted XLS / encrypted OOXML decryption (`MS-OFFCRYPTO`) | Implemented with guardrails | Modern encrypted OOXML packages (`EncryptionInfo` + `EncryptedPackage`) remain handled by `src/crypto/office-crypto.ts`. Legacy XLS `FilePass` now has `src/xls/filepass.ts` for XOR, RC4 Standard, and RC4 CryptoAPI record-payload decryption with in-file RC4/MD5/SHA-1 helpers and no npm dependency. Wrong/missing passwords surface `EncryptedFileError`. |
| old BIFF2/BIFF3/BIFF4/BIFF5 edge cases beyond BIFF8-first path | Improved | `parseBiffWorkbook` no longer requires `BoundSheet8`; raw old BIFF streams fall back to a single `Sheet1`. The worksheet parser now accepts BIFF2-style `Dimensions`, `Integer`, `Number`, `Label`, and `BoolErr` records in addition to the BIFF5/BIFF8 path. |
| XLS formula token decompilation into human-readable formulas | Implemented for core formula vocabulary; cached value path remains authoritative | `src/xls/formula.ts` decodes arithmetic/comparison operators, literals, bool/error/int/num/string constants, fixed/variable function calls, names, A1 refs/ranges, relative refs/ranges, and 3D refs/ranges. Unsupported/future tokens are kept resilient by preserving cached formula results. |
| XLSB styles/date formats from binary `styles.bin` | Implemented | `readXlsb` parses `styles.bin`, `BrtFmt`, `BrtXF`, and cell XF collections, then converts date-formatted numeric cells through the shared date utilities. |
| XLSB rich shared-string formatting runs | Implemented as public rich-text spans | `parseSharedStringsBin` now returns `XlsbSharedString` records. Rich run boundaries after `BrtSSTItem` are preserved as `Cell.richText` spans without fabricating unavailable workbook-font details. |
| charts, drawings, comments, tables, pivots, defined names, hyperlinks, document properties for XLS/XLSB parity with current XLSX reader | Implemented for values/metadata and surfaced for opaque binary parts | XLS: named ranges, hyperlinks, document properties, and CFB binary-part surfacing are present. XLSB: hyperlinks, document properties, styles, shared strings, and OPC package-part surfacing are present. Chart/drawing/comment/table/pivot parts are exposed as `binaryParts` / `packageParts` with `kind` classification so callers can preserve or inspect them even when hucre does not parse their binary internals. |
| macro/VBA project surfacing or round-trip preservation for `.xls` / `.xlsb` | Surfaced | XLS CFB streams and XLSB OPC entries classified as `vba` are exposed through `workbook.vbaProject.parts`; they are also included in `binaryParts` / `packageParts`. |

## Validation performed

The bundle includes the following validation artifacts:

- `artifacts/remaining-gap-verification.json` — machine-readable checklist status.
- `artifacts/typescript-sanity-check.txt` — TypeScript sanity check for the touched parser/crypto files using local stubs for the host hucre types.
- Existing uploaded `.xls` render artifacts remain in `artifacts/` and `examples/rendered/`.

Commands run while producing this bundle:

```sh
tsc --noEmit --module esnext --target es2022 --moduleResolution node --skipLibCheck \
  src/xls/filepass.ts src/xls/biff.ts src/xls/reader.ts src/xlsb/reader.ts

git bundle verify hucre-xls-xlsb-remaining-gaps-verified.bundle

git clone hucre-xls-xlsb-remaining-gaps-verified.bundle /tmp/hucre-bundle-checkout
```

## Scope note

This code path is still intentionally read-focused. It surfaces opaque XLS/XLSB binary package parts so they can be inspected or preserved by higher-level workflows, but it does not attempt to author chart/drawing/pivot/comment binary records from scratch.
