import { safeStorage } from 'electron'
import { ErrorCode } from '@shared/errors.js'
import { aiError } from './errors.js'

export interface SecretBox {
  seal(plain: string): string
  open(sealed: string): string
}

type SafeStorage = Pick<
  typeof safeStorage,
  'isEncryptionAvailable' | 'encryptString' | 'decryptString' | 'getSelectedStorageBackend'
>

/**
 * On Linux without a keyring Chromium falls back to `basic_text`, a fixed password anyone can read
 * from the source: a key "encrypted" that way is a key in the clear.
 */
export function createSecretBox(storage: SafeStorage = safeStorage, platform = process.platform): SecretBox {
  const assertUsable = (): void => {
    const plainText = platform === 'linux' && storage.getSelectedStorageBackend() === 'basic_text'
    if (!storage.isEncryptionAvailable() || plainText) {
      throw aiError(ErrorCode.AiSecretsUnavailable, 'errors.ai.secretsUnavailable')
    }
  }
  return {
    seal(plain) {
      assertUsable()
      return storage.encryptString(plain).toString('base64')
    },
    open(sealed) {
      assertUsable()
      return storage.decryptString(Buffer.from(sealed, 'base64'))
    },
  }
}
