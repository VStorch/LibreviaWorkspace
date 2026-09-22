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
import { BUILTIN_STYLES, type StyleSheet } from './styles.js'

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
    headerDistanceMm: 12.5,
    footerDistanceMm: 12.5,
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
  styles: BUILTIN_STYLES,
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

describe('estilos no formato interno', () => {
  const v2 = (styles?: StyleSheet): string =>
    JSON.stringify({
      format: 'sdoc',
      version: 2,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
      ...(styles === undefined ? {} : { styles }),
    })

  it('devolve os estilos gravados', () => {
    // A ida e volta precisa ser fiel também aqui: os estilos do documento são o
    // que a entrega seguinte vai usar para desenhar, e um só perdido no caminho
    // mudaria a aparência de um arquivo que ninguém editou.
    expect(parseDocument(serializeDocument(richDocument)).styles).toEqual(BUILTIN_STYLES)
  })

  it('dá os estilos embutidos ao arquivo da versão 2, que não os tinha', () => {
    // Documento antigo tem de abrir **idêntico**: os embutidos são a aparência
    // que o editor já desenhava, medida por medida.
    expect(parseDocument(v2()).styles).toEqual(BUILTIN_STYLES)
  })

  it('ignora estilos num arquivo que se declara da versão 2', () => {
    // Versão 2 não tem estilos. Um `styles` ali é remendo — provavelmente de um
    // arquivo editado à mão —, e confiar nele seria abrir o documento com uma
    // formatação que o programa que o gravou nunca conheceu.
    const forjado: StyleSheet = {
      defaults: { paragraph: {}, character: {}, paragraphStyleId: null, characterStyleId: null },
      styles: {},
    }
    expect(parseDocument(v2(forjado)).styles).toEqual(BUILTIN_STYLES)
  })

  it('dá os estilos embutidos ao arquivo da versão 1', () => {
    const version1 = JSON.stringify({
      format: 'sdoc',
      version: 1,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    expect(parseDocument(version1).styles).toEqual(BUILTIN_STYLES)
  })

  it('recusa um estilo malformado em vez de abrir o documento sem ele', () => {
    // Sem tipo, o estilo não é de parágrafo nem de caractere: não há como dizer
    // onde ele se aplica, e adivinhar é como a formatação some em silêncio.
    const quebrado = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc' },
      styles: { defaults: {}, styles: { Corpo: { id: 'Corpo', name: 'Corpo' } } },
    })
    expect(() => parseDocument(quebrado)).toThrow(/documento válido/i)
  })
})

describe('documento gravado pela versão 1 do formato', () => {
  // Na versão 1 a imagem era bloco: inserida pela barra, ficava solta entre os
  // parágrafos, ou direto dentro de uma célula. Hoje ela é inline e só existe
  // dentro de um parágrafo — e o TipTap monta o documento sem validar o schema,
  // então a imagem antiga abria "funcionando" num documento inválido.
  const image = { type: 'image', attrs: { src: 'data:image/png;base64,AAAA', width: 40 } }
  const versionOne = (doc: object): string =>
    JSON.stringify({ format: 'sdoc', version: 1, page: DEFAULT_PAGE_SETUP, doc })

  it('embrulha num parágrafo a imagem solta entre os blocos', () => {
    const parsed = parseDocument(
      versionOne({
        type: 'doc',
        content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Antes' }] }, image],
      }),
    )
    expect(parsed.doc.content).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'Antes' }] },
      { type: 'paragraph', content: [image] },
    ])
  })

  it('embrulha também a imagem solta dentro de uma célula', () => {
    const parsed = parseDocument(
      versionOne({
        type: 'doc',
        content: [
          {
            type: 'table',
            content: [{ type: 'tableRow', content: [{ type: 'tableCell', content: [image] }] }],
          },
        ],
      }),
    )
    expect(parsed.doc.content?.[0]?.content?.[0]?.content?.[0]?.content).toEqual([
      { type: 'paragraph', content: [image] },
    ])
  })

  it('não mexe na imagem que já está num parágrafo', () => {
    const doc = { type: 'doc', content: [{ type: 'paragraph', content: [image] }] }
    expect(parseDocument(versionOne(doc)).doc).toEqual(doc)
  })

  it('não migra o que já foi gravado na versão atual', () => {
    const doc = { type: 'doc', content: [image] }
    const current = JSON.stringify({ format: 'sdoc', version: SDOC_VERSION, page: DEFAULT_PAGE_SETUP, doc })
    expect(parseDocument(current).doc).toEqual(doc)
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
