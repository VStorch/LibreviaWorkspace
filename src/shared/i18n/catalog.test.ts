import { describe, expect, it } from 'vitest'
import { AREAS, MESSAGES } from './catalog/index.js'
import { LANGUAGES } from './language.js'
import { format, type Entry, type Message } from './message.js'

/**
 * O contrato do catálogo.
 *
 * Existe porque a varredura que traduz a interface é grande e mecânica, e o
 * compilador só cobre metade dela: ele exige que `Entry` tenha `pt` e `en`, e
 * não enxerga nada do que está **dentro** das duas frases. Os erros que
 * sobram são justamente os silenciosos — o `{count}` que ficou só numa das
 * línguas, o plural traduzido como singular, a chave repetida entre duas
 * áreas que o espalhamento engole sem avisar.
 *
 * Cada caso aqui é um desses, e todos já apareceram em projetos que fizeram
 * esta mesma migração.
 */

function placeholders(text: string): string[] {
  return [...text.matchAll(/\{(\w+)\}/g)].map((match) => match[1]!).sort()
}

/** Os dois textos de uma mensagem: um se for frase, dois se for plural. */
function variants(message: Message): string[] {
  return typeof message === 'string' ? [message] : [message.one, message.other]
}

describe('catálogo de traduções', () => {
  const entries = Object.entries(MESSAGES) as [string, Entry][]

  it('tem alguma coisa dentro', () => {
    // Um catálogo vazio faria todo o resto deste arquivo passar sem testar nada.
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
    // O caso real: a frase em português ganha um `{count}` e a inglesa não. A
    // tela em inglês passa a mostrar "pages" sem número nenhum, e nada falha.
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
    // Traduzir um par singular/plural como uma frase só dá "3 page" em inglês,
    // e o compilador aceita as duas formas igualmente.
    const mismatched = entries
      .filter(([, entry]) => (typeof entry.pt === 'string') !== (typeof entry.en === 'string'))
      .map(([key]) => key)

    expect(mismatched).toEqual([])
  })

  it('não repete chave entre áreas', () => {
    // O espalhamento que monta `MESSAGES` fica com a última e não avisa: a área
    // que perdeu continua compilando, e a frase dela some da tela.
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
    // Sem isto o catálogo vira um saco de nomes soltos, e a próxima pessoa não
    // sabe em qual dos arquivos procurar a frase que quer mudar.
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
    // Zero é plural nos dois idiomas — "0 páginas", "0 pages".
    expect(format(paginas, 'en', { count: 0 })).toBe('0 pages')
  })

  it('deixa à vista o buraco sem valor', () => {
    // Apagá-lo esconderia o erro; deixá-lo faz alguém consertar.
    expect(format(entry, 'pt', {})).toBe('Olá, {nome}')
  })

  it('não interpola o que veio dentro de um valor', () => {
    // Um nome de arquivo chamado "{nome}.docx" não pode disparar uma segunda
    // volta de substituição — e nome de arquivo é o que mais entra nestas frases.
    const arquivo: Entry = { pt: 'Salvo em {a}', en: 'Saved to {a}' }
    expect(format(arquivo, 'pt', { a: '{b}', b: 'nunca' })).toBe('Salvo em {b}')
  })
})
