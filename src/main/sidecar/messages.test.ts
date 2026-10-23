import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { Language, MESSAGES } from '@shared/i18n/index.js'

const { preferences } = vi.hoisted(() => ({ preferences: { language: 'pt' as Language } }))

vi.mock('../preferences.js', () => ({ editorPreferences: () => preferences }))

const { sidecarInventorySchema, sidecarMessage, sidecarNotice } = await import('./messages.js')

const SIDECAR_SOURCE = join(process.cwd(), 'sidecar', 'src')

function csharpFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) return entry.name === 'bin' || entry.name === 'obj' ? [] : csharpFiles(path)
    return entry.name.endsWith('.cs') ? [path] : []
  })
}

function keysInSidecar(): Set<string> {
  const keys = new Set<string>()
  for (const file of csharpFiles(SIDECAR_SOURCE)) {
    for (const match of readFileSync(file, 'utf8').matchAll(/"(sidecar\.[A-Za-z.]+)"/g)) keys.add(match[1]!)
  }
  return keys
}

describe('sidecar messages', () => {
  it('every key the sidecar sends has a sentence in the catalog, and the catalog has no stale key', () => {
    const used = keysInSidecar()
    const catalog = Object.keys(MESSAGES).filter((key) => key.startsWith('sidecar.'))

    expect([...used].filter((key) => !catalog.includes(key))).toEqual([])
    expect(catalog.filter((key) => !used.has(key))).toEqual([])
  })

  it('shows the sentence in the interface language', () => {
    preferences.language = Language.English
    expect(sidecarMessage('sidecar.error.notPdf')).toBe('This file is not a PDF.')

    preferences.language = Language.Portuguese
    expect(sidecarMessage('sidecar.error.notPdf')).toBe('Este arquivo não é um PDF.')
  })

  it('falls back to the generic sentence for a key this build does not know', () => {
    preferences.language = Language.Portuguese
    expect(sidecarMessage('sidecar.error.fromTheFuture')).toBe(
      'Ocorreu um erro inesperado ao processar o documento.',
    )
    expect(sidecarMessage('shell.print.defaultDocumentName')).toBe(
      'Ocorreu um erro inesperado ao processar o documento.',
    )
  })

  it('quotes the value the notice carries, in either language', () => {
    const notice = { key: 'sidecar.notice.hiddenSheet', value: 'Sales' }

    preferences.language = Language.English
    expect(sidecarNotice(notice)).toBe('the sheet "Sales" is hidden in the file and shows here like the others')

    preferences.language = Language.Portuguese
    expect(sidecarNotice(notice)).toBe('a aba "Sales" está oculta no arquivo e aparece aqui como as outras')
  })

  it('does not interpolate a value that looks like a placeholder', () => {
    preferences.language = Language.Portuguese
    expect(sidecarNotice({ key: 'sidecar.notice.textColor', value: '{value}' })).toBe('cor de texto "{value}"')
  })

  it('names an unknown notice instead of hiding it', () => {
    preferences.language = Language.English
    expect(sidecarNotice({ key: 'sidecar.notice.fromTheFuture' })).toBe(
      'content this version of the app cannot describe',
    )
  })

  it('turns the inventory into sentences, with the same cut as the IPC contract', () => {
    preferences.language = Language.Portuguese
    const inventory = sidecarInventorySchema.parse({
      invisible: Array.from({ length: 60 }, () => ({ key: 'sidecar.notice.comments' })),
      structural: [{ key: 'sidecar.notice.structureRevisions' }],
    })

    expect(inventory.invisible).toHaveLength(50)
    expect(inventory.invisible[0]).toBe('comentários')
    expect(inventory.lost).toEqual([])
    expect(inventory.structural).toEqual(['revisões de estrutura'])
  })
})
