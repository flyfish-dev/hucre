export { readXls } from "./xls/reader"
export {
  decryptOfficeEncryptedPackage,
  encryptOfficeAgilePackage,
  encryptOfficeAgilePackageParts,
  isOfficeEncryptedPackage,
} from "./crypto/office-crypto"
export type {
  AgileEncryptionOptions,
  EncryptedOfficePackageParts,
  OfficeCryptoOptions,
} from "./crypto/office-crypto"
