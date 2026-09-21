import { describe, expect, it } from 'vitest'
import type { DocumentNode } from './model.js'
import { charactersWithoutSpaces, countParagraphs } from './word-count.js'

function paragrafo(texto: string): DocumentNode {
  return texto === ''
    ? { type: 'paragraph' }
    : { type: 'paragraph', content: [{ type: 'text', text: texto }] }
}

describe('caracteres sem espaço', () => {
  it('desconta espaço, tabulação e quebra de linha', () => {
    expect(charactersWithoutSpaces('a b\tc\nd')).toBe(4)
  })

  it('desconta também o espaço inquebrável que vem do Word', () => {
    // `\u00a0` chega em data, em número e antes de unidade. Contá-lo como
    // caractere de texto daria um total maior que o do Word, sem nada na tela
    // que explicasse a diferença.
    expect(charactersWithoutSpaces('12\u00a0kg')).toBe(4)
  })

  it('conta caractere acentuado como um só', () => {
    expect(charactersWithoutSpaces('ação')).toBe(4)
  })
})

describe('parágrafos', () => {
  it('conta linha com texto e ignora linha vazia', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragrafo('Primeira'), paragrafo(''), paragrafo('Segunda')],
    }

    // Enter batido para abrir espaço não é conteúdo — e o Word também não o
    // conta.
    expect(countParagraphs(doc)).toBe(2)
  })

  it('conta título como parágrafo', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [{ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Título' }] }],
    }

    expect(countParagraphs(doc)).toBe(1)
  })

  it('não conta duas vezes o parágrafo dentro de item de lista ou de célula', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          content: [{ type: 'listItem', content: [paragrafo('Item')] }],
        },
        {
          type: 'table',
          content: [
            {
              type: 'tableRow',
              content: [{ type: 'tableCell', content: [paragrafo('Célula')] }],
            },
          ],
        },
      ],
    }

    expect(countParagraphs(doc)).toBe(2)
  })

  it('a marca de seção não é um parágrafo do texto', () => {
    // No OOXML a seção termina num `w:sectPr` guardado dentro de um parágrafo
    // vazio: é estrutura do arquivo, e o documento de evidências do corpus tem
    // seis dessas marcas no meio do texto.
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragrafo('Texto'), { type: 'paragraph', attrs: { sectionMark: true } }],
    }

    expect(countParagraphs(doc)).toBe(1)
  })
})
