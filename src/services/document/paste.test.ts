import { describe, expect, it } from 'vitest'
import { plainPasteContent } from './paste.js'

describe('colar sem formatação', () => {
  it('uma linha vira texto puro, e não parágrafo novo', () => {
    // Inserir um parágrafo aqui partiria a frase em duas — e colar o nome de uma
    // cidade dentro de um endereço é o caso comum.
    expect(plainPasteContent('São Paulo')).toEqual([{ type: 'text', text: 'São Paulo' }])
  })

  it('várias linhas viram um parágrafo por linha', () => {
    expect(plainPasteContent('um\ndois')).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'um' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'dois' }] },
    ])
  })

  it('a linha em branco do meio continua lá', () => {
    expect(plainPasteContent('um\n\ndois')).toHaveLength(3)
    expect(plainPasteContent('um\n\ndois')[1]).toEqual({ type: 'paragraph' })
  })

  it('quebra de linha do Windows e do Mac clássico contam como uma', () => {
    expect(plainPasteContent('um\r\ndois')).toHaveLength(2)
    expect(plainPasteContent('um\rdois')).toHaveLength(2)
  })

  it('a quebra do fim não deixa parágrafo vazio sobrando', () => {
    // Texto copiado de terminal, de célula de planilha e de tabela vem com ela.
    expect(plainPasteContent('só isto\n')).toEqual([{ type: 'text', text: 'só isto' }])
  })

  it('caractere de controle não entra no documento', () => {
    // Um `\u0000` vindo de arquivo binário deixaria o documento impossível de
    // salvar; `\v` é como célula de Excel separa linha.
    expect(plainPasteContent('a\u0000b')).toEqual([{ type: 'text', text: 'ab' }])
    expect(plainPasteContent('a\vb')).toHaveLength(2)
  })

  it('tabulação continua no texto', () => {
    // Ela é conteúdo: uma tabela colada como texto perde o alinhamento se a
    // tabulação sumir.
    expect(plainPasteContent('a\tb')).toEqual([{ type: 'text', text: 'a\tb' }])
  })

  it('texto vazio não insere nada', () => {
    expect(plainPasteContent('')).toEqual([])
    expect(plainPasteContent('\n')).toEqual([])
  })
})
