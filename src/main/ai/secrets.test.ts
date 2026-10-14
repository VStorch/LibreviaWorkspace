import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { createSecretBox } from './secrets.js'

function storage(backend: string, available = true) {
  return {
    isEncryptionAvailable: () => available,
    getSelectedStorageBackend: () => backend as 'basic_text',
    encryptString: (plain: string) => Buffer.from(`cifrado:${plain}`),
    decryptString: (sealed: Buffer) => sealed.toString().replace('cifrado:', ''),
  }
}

describe('cofre das chaves', () => {
  it('cifra e decifra pelo cofre do sistema', () => {
    const box = createSecretBox(storage('gnome_libsecret'), 'linux')
    const sealed = box.seal('sk-123')
    expect(sealed).not.toContain('sk-123')
    expect(box.open(sealed)).toBe('sk-123')
  })

  it('recusa guardar chave quando o Linux cai no texto puro', () => {
    const box = createSecretBox(storage('basic_text'), 'linux')
    expect(() => box.seal('sk-123')).toThrow(
      expect.objectContaining({ code: ErrorCode.AiSecretsUnavailable }),
    )
  })

  it('recusa quando não há cifra disponível', () => {
    const box = createSecretBox(storage('unknown', false), 'win32')
    expect(() => box.seal('sk-123')).toThrow(
      expect.objectContaining({ code: ErrorCode.AiSecretsUnavailable }),
    )
  })
})
