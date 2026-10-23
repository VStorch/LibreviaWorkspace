import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { Language, MESSAGES } from '@shared/i18n/index.js'

const { preferences } = vi.hoisted(() => ({ preferences: { language: 'pt' as Language } }))

vi.mock('../preferences.js', () => ({ editorPreferences: () => preferences }))

const { sidecarMessage } = await import('./messages.js')

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
})
