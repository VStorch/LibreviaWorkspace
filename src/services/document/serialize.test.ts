import { describe, expect, it } from 'vitest'
import { AppError } from '@shared/errors.js'
import {
  DEFAULT_PAGE_SETUP,
  PageOrientation,
  PageSize,
  createEmptyDocument,
  type DocumentModel,
} from './model.js'
import { SDOC_VERSION, parseDocument, serializeDocument } from './serialize.js'

const richDocument: DocumentModel = {
  page: {
    size: PageSize.Letter,
    orientation: PageOrientation.Landscape,
    margins: { top: 15, right: 15, bottom: 15, left: 15 },
    header: 'Relatório interno',
    footer: 'Página {n} de {total}',
    headerBand: null,
    footerBand: null,
    firstHeaderBand: null,
    firstFooterBand: null,
    evenHeaderBand: null,
    evenFooterBand: null,
  },
  doc: {
    type: 'doc',
    content: [
      { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Relatório' }] },
      {
        type: 'paragraph',
        attrs: { textAlign: 'justify', indent: 2 },
        content: [
          { type: 'text', marks: [{ type: 'bold' }], text: 'Negrito' },
          { type: 'text', text: ' e ' },
          {
            type: 'text',
            marks: [{ type: 'textStyle', attrs: { color: '#ff0000', fontSize: '14pt' } }],
            text: 'colorido',
          },
        ],
      },
    ],
  },
}

describe('ida e volta do formato interno', () => {
  it('preserva o documento inteiro', () => {
    // Este é o critério de aceite da Fase 2: salvar e reabrir sem perda.
    const restored = parseDocument(serializeDocument(richDocument))
    expect(restored).toEqual(richDocument)
  })

  it('preserva a configuração de página', () => {
    const restored = parseDocument(serializeDocument(richDocument))
    expect(restored.page.size).toBe(PageSize.Letter)
    expect(restored.page.orientation).toBe(PageOrientation.Landscape)
    expect(restored.page.margins).toEqual({ top: 15, right: 15, bottom: 15, left: 15 })
  })

  it('preserva cabeçalho e rodapé', () => {
    const restored = parseDocument(serializeDocument(richDocument))
    expect(restored.page.header).toBe('Relatório interno')
    expect(restored.page.footer).toBe('Página {n} de {total}')
  })

  it('abre documento gravado antes de existirem cabeçalho e rodapé', () => {
    // Compatibilidade com os arquivos da Fase 2: acrescentar campo opcional
    // não pode invalidar o que já está em disco.
    const anterior = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION,
      page: { size: 'A4', orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 } },
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
    })

    const restored = parseDocument(anterior)
    expect(restored.page.header).toBe('')
    expect(restored.page.footer).toBe('')
  })

  it('preserva um documento vazio', () => {
    const empty = createEmptyDocument()
    expect(parseDocument(serializeDocument(empty))).toEqual(empty)
  })

  it('grava a versão do formato', () => {
    expect(JSON.parse(serializeDocument(createEmptyDocument()))).toMatchObject({
      format: 'sdoc',
      version: SDOC_VERSION,
    })
  })
})

describe('leitura de arquivo problemático', () => {
  it('recusa JSON malformado com mensagem compreensível', () => {
    expect(() => parseDocument('{ isto não é json')).toThrow(AppError)
    expect(() => parseDocument('{ isto não é json')).toThrow(/corrompido|válido/i)
  })

  it('recusa JSON válido que não é um documento', () => {
    expect(() => parseDocument('{"qualquer":"coisa"}')).toThrow(/documento válido/i)
  })

  it('recusa documento de versão futura em vez de adivinhar', () => {
    const futuro = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION + 1,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc' },
    })

    expect(() => parseDocument(futuro)).toThrow(/versão mais recente/i)
  })

  it('recupera margens impossíveis usando o padrão, sem descartar o texto', () => {
    // O texto do usuário vale mais que o layout: preferimos abrir com margem
    // padrão a recusar o arquivo inteiro.
    const quebrado = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION,
      page: {
        size: 'A4',
        orientation: 'portrait',
        margins: { top: 500, right: 500, bottom: 500, left: 500 },
      },
      doc: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'importante' }] }] },
    })

    const restored = parseDocument(quebrado)
    expect(restored.page).toEqual(DEFAULT_PAGE_SETUP)
    expect(restored.doc).toMatchObject({ type: 'doc' })
  })
})
