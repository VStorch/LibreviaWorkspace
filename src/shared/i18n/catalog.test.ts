import { describe, expect, it } from 'vitest'
import { AREAS, MESSAGES } from './catalog/index.js'
import { LANGUAGES } from './language.js'
import { format, type Entry, type Message } from './message.js'

/**
 * The catalog contract beyond what the compiler checks: a `{count}` left in one language only, a
 * plural translated as a singular, a key repeated across areas that the spread swallows.
 */

function placeholders(text: string): string[] {
  return [...text.matchAll(/\{(\w+)\}/g)].map((match) => match[1]!).sort()
}

/** One text for a sentence, two for a plural. */
function variants(message: Message): string[] {
  return typeof message === 'string' ? [message] : [message.one, message.other]
}

describe('catálogo de traduções', () => {
  const entries = Object.entries(MESSAGES) as [string, Entry][]

  it('tem alguma coisa dentro', () => {
    // An empty catalog would make the rest of this file pass without testing anything.
    expect(entries.length).toBeGreaterThan(0)
  })

  it('não deixa nenhuma frase vazia', () => {
    const empty = entries.flatMap(([key, entry]) =>
      LANGUAGES.flatMap((language) =>
        variants(entry[language])
          .filter((text) => text.trim() === '')
          .map(() => `${key}.${language}`),
      ),
    )

    expect(empty).toEqual([])
  })

  it('usa os mesmos buracos nos dois idiomas', () => {
    // The real case: the Portuguese sentence gains a `{count}` and the English one does not. The
    // English UI then shows "pages" with no number, and nothing fails.
    const mismatched = entries
      .filter(([, entry]) => {
        const pt = variants(entry.pt).flatMap(placeholders)
        const en = variants(entry.en).flatMap(placeholders)
        return new Set(pt).size !== new Set(en).size || pt.some((name) => !en.includes(name))
      })
      .map(([key]) => key)

    expect(mismatched).toEqual([])
  })

  it('é plural nos dois idiomas ou em nenhum', () => {
    // Translating a singular/plural pair as a single sentence gives "3 page" in English, and the
    // compiler accepts both shapes.
    const mismatched = entries
      .filter(([, entry]) => (typeof entry.pt === 'string') !== (typeof entry.en === 'string'))
      .map(([key]) => key)

    expect(mismatched).toEqual([])
  })

  it('não repete chave entre áreas', () => {
    // The spread that builds `MESSAGES` keeps the last one silently: the losing area still
    // compiles, and its sentence disappears from the screen.
    const seen = new Map<string, string>()
    const clashes: string[] = []

    for (const [area, catalog] of Object.entries(AREAS)) {
      for (const key of Object.keys(catalog)) {
        const owner = seen.get(key)
        if (owner !== undefined) clashes.push(`${key}: ${owner} e ${area}`)
        else seen.set(key, area)
      }
    }

    expect(clashes).toEqual([])
  })

  it('prefixa cada chave com a área em que mora', () => {
    // Otherwise the catalog becomes a bag of loose names, and nobody knows which file holds the
    // sentence.
    const misplaced = Object.entries(AREAS).flatMap(([area, catalog]) =>
      Object.keys(catalog).filter((key) => !key.startsWith(`${area}.`)),
    )

    expect(misplaced).toEqual([])
  })
})

describe('format', () => {
  const entry: Entry = { pt: 'Olá, {nome}', en: 'Hello, {nome}' }

  it('escolhe o idioma', () => {
    expect(format(entry, 'pt', { nome: 'Ana' })).toBe('Olá, Ana')
    expect(format(entry, 'en', { nome: 'Ana' })).toBe('Hello, Ana')
  })

  it('escolhe singular e plural pelo count', () => {
    const paginas: Entry = {
      pt: { one: '{count} página', other: '{count} páginas' },
      en: { one: '{count} page', other: '{count} pages' },
    }

    expect(format(paginas, 'pt', { count: 1 })).toBe('1 página')
    expect(format(paginas, 'pt', { count: 2 })).toBe('2 páginas')
    expect(format(paginas, 'en', { count: 0 })).toBe('0 pages')
  })

  it('deixa à vista o buraco sem valor', () => {
    // Dropping it would hide the mistake; keeping it gets someone to fix it.
    expect(format(entry, 'pt', {})).toBe('Olá, {nome}')
  })

  it('não interpola o que veio dentro de um valor', () => {
    // A file named "{nome}.docx" must not trigger a second substitution pass, and file names are
    // what these sentences carry most.
    const arquivo: Entry = { pt: 'Salvo em {a}', en: 'Saved to {a}' }
    expect(format(arquivo, 'pt', { a: '{b}', b: 'nunca' })).toBe('Salvo em {b}')
  })
})
