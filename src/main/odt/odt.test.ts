/**
 * The ODT `odt.create` + `odt.save` write, read back by the published sidecar (skipped when it is
 * not published): the Phase 1 scope comes back as the same model. Text in a list or a cell carries
 * its effective formatting, as the DOCX reader flattens it there. `Normal` becomes LibreOffice's
 * `Standard`, which the reader leaves unsaid, as the DOCX reader leaves `Normal`. Then saved back
 * into the same package, and to other formats.
 */

import { access, constants, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, type DocumentModel, type DocumentNode } from '@services/document/model.js'
import { BUILTIN_STYLES, StyleType, type StyleSheet } from '@services/document/styles.js'
import { SidecarClient } from '../sidecar/client.js'
import { sidecarPathIn } from '../sidecar/locate.js'
import { saveDocx, sdocContentOf, withoutOids } from '../docx/index.js'
import { t } from '../i18n.js'
import {
  followOriginal,
  forgetOpenedOriginal,
  isCreatedOriginal,
  OriginalFormat,
} from '../opened-original.js'
import { decodeOdt, exportOdt, openOdt, saveOdt } from './index.js'

const client = new SidecarClient(() => Promise.resolve(sidecarPathIn(process.cwd())))
const GENERATOR = 'Librevia/0.0.0'

afterAll(() => client.dispose())

const published = await access(sidecarPathIn(process.cwd()), constants.X_OK).then(
  () => true,
  () => false,
)

/** A 2 × 3 pixel PNG, header only: enough to measure. */
const PNG = (() => {
  const bytes = new Uint8Array(33)
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52])
  bytes.set([0, 0, 0, 2, 0, 0, 0, 3], 16)
  return `data:image/png;base64,${btoa(String.fromCharCode(...bytes))}`
})()

/** ODF states line spacing in whole percent: 1.08, not Word's 1.0792. */
const STYLES: StyleSheet = {
  defaults: {
    ...BUILTIN_STYLES.defaults,
    paragraph: { spaceAfter: 8, lineSpacing: { kind: 'multiple', factor: 1.08 } },
  },
  styles: {
    ...BUILTIN_STYLES.styles,
    Destaque: {
      id: 'Destaque',
      name: 'Destaque',
      type: StyleType.Character,
      qFormat: true,
      hidden: false,
      custom: true,
      character: { bold: true },
    },
  },
}

type Marks = DocumentNode['marks']

const text = (value: string, marks?: Marks): DocumentNode =>
  marks === undefined ? { type: 'text', text: value } : { type: 'text', text: value, marks }

/** The block where the style CSS does not reach: Normal over the defaults, Calibri 11 at 1.08. */
const flat = (value: string): DocumentNode => ({
  type: 'paragraph',
  attrs: {
    styleId: 'Normal',
    spaceBefore: 0,
    spaceAfter: 8,
    lineHeight: '1.3184',
    fontFamily: 'Calibri',
    fontSize: '11pt',
  },
  content: [text(value)],
})

const item = (...content: DocumentNode[]): DocumentNode => ({ type: 'listItem', content })

const cell = (value: string, attrs: Record<string, unknown>): DocumentNode => ({
  type: 'tableCell',
  attrs: { colspan: 1, rowspan: 1, ...attrs },
  content: [flat(value)],
})

const level = (fmt: string, value: string, start: number, indentMm: number) => ({
  fmt,
  text: value,
  start,
  indentMm,
  hangingMm: 6.35,
})

const paragraph = (): DocumentNode => ({
  type: 'paragraph',
  attrs: {
    styleId: 'Normal',
    textAlign: 'center',
    spaceBefore: 6,
    spaceAfter: 12,
    lineHeight: '2.2998',
    fontFamily: 'Arial',
    keepNext: true,
  },
  content: [
    text('Olá  '),
    text('negrito', [{ type: 'bold' }]),
    text(' '),
    text('itálico', [{ type: 'italic' }]),
    text(' sub', [{ type: 'strike' }, { type: 'underline' }]),
    text(' cor', [
      { type: 'textStyle', attrs: { color: '#c00000', fontSize: '14pt', fontFamily: 'Courier New' } },
    ]),
    text(' realce', [{ type: 'highlight', attrs: { color: '#ffff00' } }]),
    text(' link', [{ type: 'link', attrs: { href: 'https://example.com/?a=1&b=2' } }]),
    text(' destaque', [{ type: 'bold' }, { type: 'charStyle', attrs: { styleId: 'Destaque' } }]),
    text(' alto', [{ type: 'superscript' }]),
    { type: 'hardBreak' },
    { type: 'image', attrs: { src: PNG, alt: 'figura', width: 96, height: 48 } },
  ],
})

const lists = (): DocumentNode[] => [
  {
    type: 'orderedList',
    attrs: {
      start: 1,
      numbering: {
        key: 'odt1',
        levels: [level('decimal', '%1.', 3, 12.7), level('lowerLetter', '%1.%2)', 1, 25.4)],
      },
    },
    content: [
      item(flat('três')),
      item(flat('quatro'), { type: 'orderedList', attrs: { start: 1 }, content: [item(flat('sub'))] }),
    ],
  },
  {
    type: 'bulletList',
    attrs: { numbering: { key: 'odt2', levels: [level('bullet', '•', 1, 12.7)] } },
    content: [item(flat('marcador'))],
  },
]

const table = (): DocumentNode => ({
  type: 'table',
  content: [
    {
      type: 'tableRow',
      content: [cell('A', { colspan: 2, colwidth: [96, 96] }), cell('B', { rowspan: 2, colwidth: [96] })],
    },
    {
      type: 'tableRow',
      content: [
        cell('C', { colwidth: [96] }),
        cell('D', { colwidth: [96], shading: '#eeeeee', borders: 'top:double,1.5,#ff0000' }),
      ],
    },
  ],
})

/**
 * The export test's rich model without what the reader leaves for later: notes, comments, fields,
 * bookmarks and sections.
 */
const MODEL: DocumentModel = {
  page: {
    ...DEFAULT_PAGE_SETUP,
    header: 'Relatório — página {n} de {total}',
    margins: { top: 30, right: 20, bottom: 20, left: 30 },
  },
  styles: STYLES,
  properties: {
    title: 'Relatório <anual> & "final"',
    subject: 'Contas',
    creator: 'Ana',
    lastModifiedBy: 'Bia',
    keywords: 'contas; balanço',
    created: '2026-01-02T03:04:05Z',
    modified: '2026-02-03T04:05:06Z',
  },
  doc: {
    type: 'doc',
    content: [
      {
        type: 'heading',
        attrs: { level: 1, styleId: 'Heading1' },
        content: [text('Título <script>&')],
      },
      paragraph(),
      ...lists(),
      { type: 'pageBreak' },
      table(),
      { type: 'paragraph', attrs: { styleId: 'Normal' }, content: [text('Fim')] },
    ],
  },
}

/** A document without an ODT original, through the minimal package. */
const created = async (): Promise<Buffer> =>
  Buffer.from(await exportOdt(client, sdocContentOf(MODEL), GENERATOR))

describe.skipIf(!published)('odt.open', () => {
  it('reads back what odt.create and odt.save write', async () => {
    const opened = await decodeOdt(client, '/tmp/export.odt', await created())
    const model = JSON.parse(opened.content) as DocumentModel

    expect(opened.inventory).toEqual({ invisible: [], lost: [], structural: [] })
    expect(withoutOids(model.doc)).toEqual(withoutNormal(MODEL.doc))
    expect(model.properties).toEqual(MODEL.properties)
    expect(model.page).toMatchObject({
      size: 'A4',
      orientation: 'portrait',
      margins: MODEL.page.margins,
      header: MODEL.page.header,
      footer: '',
      headerDistanceMm: 12.5,
    })
  })

  it('reads back the style sheet', async () => {
    const opened = await decodeOdt(client, '/tmp/export.odt', await created())
    const { styles } = JSON.parse(opened.content) as DocumentModel

    expect(styles.defaults).toMatchObject({
      paragraph: { ...STYLES.defaults.paragraph, widowControl: true },
      character: STYLES.defaults.character,
    })
    // ODF has no gallery, hiding or priority: only what the style draws comes back.
    for (const { id, name, type, basedOn, next, paragraph, character } of Object.values(STYLES.styles)) {
      const declared = JSON.parse(JSON.stringify({ id, name, type, basedOn, next, paragraph, character }))
      expect(styles.styles[id]).toMatchObject(declared)
    }
  })
})

describe.skipIf(!published)('odt.save', () => {
  let directory = ''
  let path = ''

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-odt-'))
    path = join(directory, 'relatório.odt')
    await writeFile(path, await created())
  })

  afterAll(() => rm(directory, { recursive: true, force: true }))

  /** The last block's text replaced, as the editor would send it. */
  function editLast(content: string, value: string): string {
    const model = JSON.parse(content) as DocumentModel
    const blocks = [...(model.doc.content ?? [])]
    const last = blocks.pop()
    if (last !== undefined) blocks.push({ ...last, content: [text(value)] })
    return JSON.stringify({ ...model, doc: { ...model.doc, content: blocks } })
  }

  const lastText = (content: string): string | undefined =>
    (JSON.parse(content) as DocumentModel).doc.content?.at(-1)?.content?.[0]?.text

  it('writes the edit back into the opened .odt and reads it again', async () => {
    const opened = await openOdt(client, path)
    const saved = await saveOdt(client, editLast(opened.content, 'Fim editado'), {
      origin: path,
      generator: GENERATOR,
    })
    expect(saved.created).toBe(false)
    await writeFile(path, saved.bytes)

    const reopened = await openOdt(client, path)
    expect(lastText(reopened.content)).toBe('Fim editado')
    expect(saved.inventory.lost).toEqual([])
    expect(withoutOids((JSON.parse(reopened.content) as DocumentModel).doc.content?.[0])).toEqual(
      withoutNormal(MODEL.doc).content?.[0],
    )
  })

  it('saves an opened .odt as .docx from the model, without the original package loss', async () => {
    const opened = await openOdt(client, path)
    const saved = await saveDocx(client, opened.content, {
      origin: path,
      destination: join(directory, 'a.docx'),
    })

    expect(saved.inventory.lost).not.toContain(t('errors.docx.originPackage'))
    expect(saved.bytes.length).toBeGreaterThan(0)
  })

  it('saves a document without an ODT original into the minimal package, rebuilt on each save', async () => {
    forgetOpenedOriginal()
    const content = editLast(JSON.stringify({ ...MODEL, doc: withOids(MODEL.doc) }), 'Novo')
    const target = { origin: null, generator: GENERATOR }

    const saved = await saveOdt(client, content, target)
    expect(saved.created).toBe(true)
    expect(saved.inventory.lost).toEqual([])
    const destination = join(directory, 'novo.odt')
    await writeFile(destination, saved.bytes)
    followOriginal(null, destination, { original: saved.original, format: OriginalFormat.Odt, created: true })
    expect(isCreatedOriginal(destination)).toBe(true)

    const again = await saveOdt(client, content, { ...target, origin: destination })
    expect(again.created).toBe(true)
    const reopened = await openOdt(client, destination)
    expect(lastText(reopened.content)).toBe('Novo')
    expect(isCreatedOriginal(destination)).toBe(false)
  })
})

/** The default paragraph style, as the reader gives it: unsaid. */
function withoutNormal(node: DocumentNode): DocumentNode {
  const { attrs, content, ...rest } = node
  const { styleId, ...others } = attrs ?? {}
  const kept = styleId === 'Normal' ? others : attrs
  return {
    ...rest,
    ...(kept === undefined || Object.keys(kept).length === 0 ? {} : { attrs: kept }),
    ...(content === undefined ? {} : { content: content.map(withoutNormal) }),
  }
}

/** As a DOCX gives them: they mean nothing in another package. */
function withOids(doc: DocumentNode): DocumentNode {
  return {
    ...doc,
    content: (doc.content ?? []).map((block, index) => ({
      ...block,
      attrs: { ...block.attrs, oid: `docx-${index}` },
    })),
  }
}
