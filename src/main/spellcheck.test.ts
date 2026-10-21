import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterAll, describe, expect, it, vi } from 'vitest'
import { DICTIONARY_FOLDER, dictionaryFileName, hasBdictSignature } from '@services/spell/dictionary.js'

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..', '..')

/**
 * A fake profile: `installBundledDictionary` only needs `app.getPath('userData')` and a false
 * `app.isPackaged`.
 */
const perfil = mkdtempSync(join(tmpdir(), 'librevia-spell-'))

vi.mock('electron', () => ({
  app: { isPackaged: false, getPath: () => perfil },
  protocol: { handle: () => {} },
}))

const { installBundledDictionary } = await import('./spellcheck.js')

afterAll(() => {
  rmSync(perfil, { recursive: true, force: true })
})

describe('bundled dictionary', () => {
  it('exists in resources/dictionaries and is in Chromium format', () => {
    // Same reason as the fonts test (`src/main/fonts.test.ts`): without it the failure is silent
    // and only shows on the installer's machine, which has no dictionary in the profile and,
    // offline, ends up without a spellchecker.
    const arquivo = join(raiz, 'resources', DICTIONARY_FOLDER.toLowerCase(), dictionaryFileName())
    const conteudo = readFileSync(arquivo)

    expect(hasBdictSignature(conteudo)).toBe(true)
    // A full Portuguese dictionary is a few megabytes. The floor catches a file that became a Git
    // LFS pointer or an interrupted download.
    expect(conteudo.byteLength).toBeGreaterThan(1_000_000)
  })
})

describe('dictionary installation in the profile', () => {
  const instalado = join(perfil, DICTIONARY_FOLDER, dictionaryFileName())

  it('copies the bundled dictionary on first run', () => {
    expect(installBundledDictionary()).toBe(true)
    expect(hasBdictSignature(readFileSync(instalado))).toBe(true)
  })

  it('repairs the corrupt file in the same run', () => {
    // Chromium deletes a corrupt `.bdic` and tries to download: offline, spelling disappears
    // without warning. The signature check restores it right away.
    writeFileSync(instalado, 'lixo')

    expect(installBundledDictionary()).toBe(true)
    expect(hasBdictSignature(readFileSync(instalado))).toBe(true)
  })
})
