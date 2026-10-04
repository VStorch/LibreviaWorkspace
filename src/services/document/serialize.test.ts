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
import { BUILTIN_STYLES, LEGACY_STYLES, type StyleSheet } from './styles.js'

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
    // Campo opcional novo não invalida o que está em disco.
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
    // Os estilos do documento atravessam fiéis: um perdido mudaria a aparência.
    expect(parseDocument(serializeDocument(richDocument)).styles).toEqual(BUILTIN_STYLES)
  })

  it('marca como achatados os blocos de um arquivo anterior à versão 4', () => {
    // Antes da versão 4 cada bloco levava a formatação efetiva, e a gravação em DOCX
    // compara com uma leitura achatada.
    expect(parseDocument(v2()).flattened).toBe(true)
    expect(parseDocument(serializeDocument(richDocument)).flattened).toBeUndefined()
    const flat = parseDocument(serializeDocument({ ...richDocument, flattened: true }))
    expect(flat.flattened).toBe(true)
  })

  it('marca como anterior às referências o arquivo de antes da versão 5', () => {
    // Antes da versão 5 os nós não trazem marcador nem campo.
    const v4 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 4 })
    expect(parseDocument(v4).beforeReferences).toBe(true)
    expect(parseDocument(v4).flattened).toBeUndefined()
    expect(parseDocument(serializeDocument(richDocument)).beforeReferences).toBeUndefined()
    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeReferences: true }))
    expect(legacy.beforeReferences).toBe(true)
  })

  it('marca como anterior às seções o arquivo de antes da versão 6, e leva as seções na ida e volta', () => {
    // Antes da versão 6 não há `sections` nem `sectionBreak`: a página é a do documento inteiro.
    const v5 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 5 })
    expect(parseDocument(v5).beforeSections).toBe(true)
    expect(parseDocument(v5).sections).toBeUndefined()
    expect(parseDocument(serializeDocument(richDocument)).beforeSections).toBeUndefined()

    const landscape = {
      ...DEFAULT_PAGE_SETUP,
      id: 's1',
      orientation: 'landscape' as const,
      start: 'continuous' as const,
    }
    const reopened = parseDocument(serializeDocument({ ...richDocument, sections: [landscape] }))
    expect(reopened.sections).toEqual([landscape])
    expect(reopened.beforeSections).toBeUndefined()
  })

  it('marca como anterior aos comentários o arquivo da versão 6, e leva os comentários na ida e volta', () => {
    // Antes da versão 7 os nós não trazem as pontas da âncora.
    const v6 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 6 })
    expect(parseDocument(v6).beforeComments).toBe(true)
    expect(parseDocument(v6).comments).toBeUndefined()
    expect(parseDocument(serializeDocument(richDocument)).beforeComments).toBeUndefined()

    const thread = [
      {
        id: '0',
        author: 'Ana',
        initials: 'A',
        date: '2026-03-02T10:00:00Z',
        paragraphs: ['Conferir.'],
        done: false,
      },
      { id: '1', parentId: '0', author: 'Bruno', date: '', paragraphs: ['Ok.'], done: true, rich: true },
    ]
    const reopened = parseDocument(serializeDocument({ ...richDocument, comments: thread }))
    expect(reopened.comments).toEqual(thread)
    expect(reopened.beforeComments).toBeUndefined()
    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeComments: true }))
    expect(legacy.beforeComments).toBe(true)
  })

  it('marca como anterior às revisões o arquivo da versão 7, e leva o interruptor na ida e volta', () => {
    // Antes do controle de alterações, sem marcas de revisão.
    const v7 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 7 })
    expect(parseDocument(v7).beforeRevisions).toBe(true)
    expect(parseDocument(v7).beforeComments).toBeUndefined()
    expect(parseDocument(serializeDocument(richDocument)).beforeRevisions).toBeUndefined()
    expect(JSON.parse(serializeDocument(richDocument)).version).toBe(SDOC_VERSION)

    const tracked = parseDocument(serializeDocument({ ...richDocument, trackChanges: true }))
    expect(tracked.trackChanges).toBe(true)
    expect(parseDocument(serializeDocument(richDocument)).trackChanges).toBeUndefined()
    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeRevisions: true }))
    expect(legacy.beforeRevisions).toBe(true)
  })

  it('marca como anterior às notas o arquivo da versão 8, e leva a numeração na ida e volta', () => {
    // Antes das notas, sem `noteRef`.
    const v8 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 8 })
    expect(parseDocument(v8).beforeNotes).toBe(true)
    expect(parseDocument(v8).beforeRevisions).toBeUndefined()
    expect(JSON.parse(serializeDocument(richDocument)).version).toBe(SDOC_VERSION)
    expect(parseDocument(serializeDocument(richDocument)).beforeNotes).toBeUndefined()

    const notes = { footnotePr: { numFmt: 'lowerRoman', start: 3 }, endnotePr: { pos: 'docEnd' } }
    expect(parseDocument(serializeDocument({ ...richDocument, notes })).notes).toEqual(notes)
    expect(parseDocument(serializeDocument(richDocument)).notes).toBeUndefined()
    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeNotes: true }))
    expect(legacy.beforeNotes).toBe(true)
  })

  it('marca como anterior às equações o arquivo da versão 10, e o atual não', () => {
    // Antes das equações não há nó `math`, e editar o parágrafo perde a equação: a
    // gravação em DOCX declara isso.
    const v10 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 10 })
    expect(parseDocument(v10).beforeMath).toBe(true)
    expect(parseDocument(v10).beforeNotes).toBeUndefined()
    expect(JSON.parse(serializeDocument(richDocument)).version).toBe(11)
    expect(parseDocument(serializeDocument(richDocument)).beforeMath).toBeUndefined()

    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeMath: true }))
    expect(legacy.beforeMath).toBe(true)

    // O nó atravessa o `.sdoc` como veio: o OMML é o que volta ao arquivo.
    const omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/>'
    const math = { type: 'math', attrs: { omml, mathml: '<math></math>', latex: '', display: false } }
    const doc = { type: 'doc', content: [{ type: 'paragraph', content: [math] }] }
    expect(parseDocument(serializeDocument({ ...richDocument, doc })).doc).toEqual(doc)
  })

  it('leva as propriedades na ida e volta, e o arquivo da versão 9 abre sem elas e sem marca', () => {
    // Antes das propriedades: a ausência deixa `docProps/` como está no arquivo.
    const v9 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 9 })
    const legacy = parseDocument(v9)
    expect(legacy.properties).toBeUndefined()
    expect(Object.keys(legacy).some((key) => key.startsWith('before') && key.includes('Propert'))).toBe(false)
    expect(JSON.parse(serializeDocument(richDocument)).version).toBe(SDOC_VERSION)

    const properties = {
      title: 'Relatório',
      keywords: 'a; b',
      creator: 'Ana',
      created: '2026-01-02T03:04:05Z',
      totalTime: 12,
      company: '',
    }
    expect(parseDocument(serializeDocument({ ...richDocument, properties })).properties).toEqual(properties)
    expect(parseDocument(serializeDocument(richDocument)).properties).toBeUndefined()
  })

  it('recusa propriedades malformadas', () => {
    const broken = JSON.stringify({
      ...JSON.parse(serializeDocument(richDocument)),
      properties: { title: 42 },
    })
    expect(() => parseDocument(broken)).toThrow(AppError)
    const long = JSON.stringify({
      ...JSON.parse(serializeDocument(richDocument)),
      properties: { totalTime: -1 },
    })
    expect(() => parseDocument(long)).toThrow(AppError)
  })

  it('dá os estilos embutidos ao arquivo da versão 2, que não os tinha', () => {
    // Documento antigo abre idêntico, medida por medida.
    expect(parseDocument(v2()).styles).toEqual(LEGACY_STYLES)
  })

  it('ignora estilos num arquivo que se declara da versão 2', () => {
    // A versão 2 não tem estilos: um `styles` ali é remendo, e não se confia nele.
    const forjado: StyleSheet = {
      defaults: { paragraph: {}, character: {}, paragraphStyleId: null, characterStyleId: null },
      styles: {},
    }
    expect(parseDocument(v2(forjado)).styles).toEqual(LEGACY_STYLES)
  })

  it('dá os estilos embutidos ao arquivo da versão 1', () => {
    const version1 = JSON.stringify({
      format: 'sdoc',
      version: 1,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    expect(parseDocument(version1).styles).toEqual(LEGACY_STYLES)
  })

  it('recusa um estilo malformado em vez de abrir o documento sem ele', () => {
    // Sem tipo, não há onde o estilo se aplique.
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
  // Na versão 1 a imagem era bloco, solta entre parágrafos ou na célula; hoje é
  // inline, e o TipTap monta o documento sem validar o schema.
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
    // Melhor margem padrão que recusar o arquivo.
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
