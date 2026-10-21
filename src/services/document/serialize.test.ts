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

describe('internal format round trip', () => {
  it('preserves the whole document', () => {
    const restored = parseDocument(serializeDocument(richDocument))
    expect(restored).toEqual(richDocument)
  })

  it('preserves the page setup', () => {
    const restored = parseDocument(serializeDocument(richDocument))
    expect(restored.page.size).toBe(PageSize.Letter)
    expect(restored.page.orientation).toBe(PageOrientation.Landscape)
    expect(restored.page.margins).toEqual({ top: 15, right: 15, bottom: 15, left: 15 })
  })

  it('preserves header and footer', () => {
    const restored = parseDocument(serializeDocument(richDocument))
    expect(restored.page.header).toBe('Relatório interno')
    expect(restored.page.footer).toBe('Página {n} de {total}')
  })

  it('opens a document saved before header and footer existed', () => {
    // A new optional field does not invalidate what is on disk.
    const previous = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION,
      page: { size: 'A4', orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 } },
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
    })

    const restored = parseDocument(previous)
    expect(restored.page.header).toBe('')
    expect(restored.page.footer).toBe('')
  })

  it('preserves an empty document', () => {
    const empty = createEmptyDocument()
    expect(parseDocument(serializeDocument(empty))).toEqual(empty)
  })

  it('writes the format version', () => {
    expect(JSON.parse(serializeDocument(createEmptyDocument()))).toMatchObject({
      format: 'sdoc',
      version: SDOC_VERSION,
    })
  })
})

describe('styles in the internal format', () => {
  const v2 = (styles?: StyleSheet): string =>
    JSON.stringify({
      format: 'sdoc',
      version: 2,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
      ...(styles === undefined ? {} : { styles }),
    })

  it('returns the saved styles', () => {
    // The document styles cross faithfully: a lost one would change the look.
    expect(parseDocument(serializeDocument(richDocument)).styles).toEqual(BUILTIN_STYLES)
  })

  it('marks the blocks of a file older than version 4 as flattened', () => {
    // Before version 4 each block carried the effective formatting, and saving to DOCX compares
    // with a flattened reading.
    expect(parseDocument(v2()).flattened).toBe(true)
    expect(parseDocument(serializeDocument(richDocument)).flattened).toBeUndefined()
    const flat = parseDocument(serializeDocument({ ...richDocument, flattened: true }))
    expect(flat.flattened).toBe(true)
  })

  it('marks a file older than version 5 as predating references', () => {
    // Before version 5 the nodes carry no bookmarks or fields.
    const v4 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 4 })
    expect(parseDocument(v4).beforeReferences).toBe(true)
    expect(parseDocument(v4).flattened).toBeUndefined()
    expect(parseDocument(serializeDocument(richDocument)).beforeReferences).toBeUndefined()
    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeReferences: true }))
    expect(legacy.beforeReferences).toBe(true)
  })

  it('marks a file older than version 6 as predating sections, and carries sections in the round trip', () => {
    // Before version 6 there are no `sections` or `sectionBreak`: the page is the whole document's.
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

  it('marks a version 6 file as predating comments, and carries comments in the round trip', () => {
    // Before version 7 the nodes carry no anchor ends.
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

  it('marks a version 7 file as predating revisions, and carries the switch in the round trip', () => {
    // Before track changes, no revision marks.
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

  it('marks a version 8 file as predating notes, and carries the numbering in the round trip', () => {
    // Before notes, no `noteRef`.
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

  it('keeps the request not to write names into the properties', () => {
    const properties = { creator: '', removePersonalInformation: true }
    expect(parseDocument(serializeDocument({ ...richDocument, properties })).properties).toEqual(properties)
  })

  it('marks a version 10 file as predating equations, and the current one not', () => {
    // Before equations there is no `math` node, and editing the paragraph loses the equation:
    // saving to DOCX says so.
    const v10 = JSON.stringify({ ...JSON.parse(v2(BUILTIN_STYLES)), version: 10 })
    expect(parseDocument(v10).beforeMath).toBe(true)
    expect(parseDocument(v10).beforeNotes).toBeUndefined()
    expect(JSON.parse(serializeDocument(richDocument)).version).toBe(SDOC_VERSION)
    expect(parseDocument(serializeDocument(richDocument)).beforeMath).toBeUndefined()

    const legacy = parseDocument(serializeDocument({ ...richDocument, beforeMath: true }))
    expect(legacy.beforeMath).toBe(true)

    // The node crosses the `.sdoc` as it came: the OMML is what goes back to the file.
    const omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/>'
    const math = { type: 'math', attrs: { omml, mathml: '<math></math>', latex: '', display: false } }
    const doc = { type: 'doc', content: [{ type: 'paragraph', content: [math] }] }
    expect(parseDocument(serializeDocument({ ...richDocument, doc })).doc).toEqual(doc)
  })

  it('carries properties in the round trip, and a version 9 file opens without them and without a mark', () => {
    // Before properties: their absence leaves `docProps/` as it is in the file.
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

  it('rejects malformed properties', () => {
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

  it('gives the built-in styles to a version 2 file, which had none', () => {
    // An old document opens identical, measure by measure.
    expect(parseDocument(v2()).styles).toEqual(LEGACY_STYLES)
  })

  it('ignores styles in a file that declares itself version 2', () => {
    // Version 2 has no styles: a `styles` there is a patch, and it is not trusted.
    const forged: StyleSheet = {
      defaults: { paragraph: {}, character: {}, paragraphStyleId: null, characterStyleId: null },
      styles: {},
    }
    expect(parseDocument(v2(forged)).styles).toEqual(LEGACY_STYLES)
  })

  it('gives the built-in styles to a version 1 file', () => {
    const version1 = JSON.stringify({
      format: 'sdoc',
      version: 1,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc', content: [{ type: 'paragraph' }] },
    })
    expect(parseDocument(version1).styles).toEqual(LEGACY_STYLES)
  })

  it('rejects a malformed style instead of opening the document without it', () => {
    // Without a type, there is nowhere for the style to apply.
    const broken = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc' },
      styles: { defaults: {}, styles: { Corpo: { id: 'Corpo', name: 'Corpo' } } },
    })
    expect(() => parseDocument(broken)).toThrow(/documento válido/i)
  })
})

describe('document saved by version 1 of the format', () => {
  // In version 1 the image was a block, loose between paragraphs or in a cell; today it is inline,
  // and Tiptap builds the document without validating the schema.
  const image = { type: 'image', attrs: { src: 'data:image/png;base64,AAAA', width: 40 } }
  const versionOne = (doc: object): string =>
    JSON.stringify({ format: 'sdoc', version: 1, page: DEFAULT_PAGE_SETUP, doc })

  it('wraps the image loose between blocks in a paragraph', () => {
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

  it('also wraps the image loose inside a cell', () => {
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

  it('leaves alone an image already in a paragraph', () => {
    const doc = { type: 'doc', content: [{ type: 'paragraph', content: [image] }] }
    expect(parseDocument(versionOne(doc)).doc).toEqual(doc)
  })

  it('does not migrate what was already saved in the current version', () => {
    const doc = { type: 'doc', content: [image] }
    const current = JSON.stringify({ format: 'sdoc', version: SDOC_VERSION, page: DEFAULT_PAGE_SETUP, doc })
    expect(parseDocument(current).doc).toEqual(doc)
  })
})

describe('reading a problematic file', () => {
  it('rejects malformed JSON with an understandable message', () => {
    expect(() => parseDocument('{ isto não é json')).toThrow(AppError)
    expect(() => parseDocument('{ isto não é json')).toThrow(/corrompido|válido/i)
  })

  it('rejects valid JSON that is not a document', () => {
    expect(() => parseDocument('{"qualquer":"coisa"}')).toThrow(/documento válido/i)
  })

  it('rejects a document from a future version instead of guessing', () => {
    const future = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION + 1,
      page: DEFAULT_PAGE_SETUP,
      doc: { type: 'doc' },
    })

    expect(() => parseDocument(future)).toThrow(/versão mais recente/i)
  })

  it('recovers impossible margins using the default, without discarding the text', () => {
    // Better the default margin than refusing the file.
    const broken = JSON.stringify({
      format: 'sdoc',
      version: SDOC_VERSION,
      page: {
        size: 'A4',
        orientation: 'portrait',
        margins: { top: 500, right: 500, bottom: 500, left: 500 },
      },
      doc: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'importante' }] }] },
    })

    const restored = parseDocument(broken)
    expect(restored.page).toEqual(DEFAULT_PAGE_SETUP)
    expect(restored.doc).toMatchObject({ type: 'doc' })
  })
})
