/**
 * The fingerprint contract, with both real sides: the document is read by the published sidecar,
 * goes through the schema built with the editor's real extensions and returns to the sidecar. The
 * schema rewrites the model (materializes attributes, merges texts, sorts marks, drops the
 * unknown), and a difference that only shows at this seam regenerates the document without failing
 * anywhere.
 *
 * Two criteria: zero rewritten blocks, and both trees equal in what the fingerprint compares,
 * because the count alone would let through what the schema and the fingerprint both ignored.
 *
 * Without Tiptap's `Editor`, which needs a DOM: what it does with `content` is `Node.fromJSON` over
 * the schema, and `getJSON()` is `toJSON()`. The extensions come through a path in a variable
 * because `tsconfig.node.json` does not know `src/renderer`.
 */

import { access, constants } from 'node:fs/promises'
import { promisify } from 'node:util'
import { inflateRaw } from 'node:zlib'
import { afterAll, describe, expect, it } from 'vitest'
import { getSchema, type Extensions } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  docxWithBulletList,
  docxWithComment,
  docxWithCommentThread,
  docxWithDirectOverStyles,
  docxWithDescribedImage,
  docxWithEquations,
  docxWithFootnote,
  docxWithHeaderGrid,
  docxWithMultilevelList,
  docxWithNamedStyles,
  docxWithSpacingOnBothSides,
  docxWithStretchedImage,
  docxWithStyledCells,
  docxWithTable,
  docxWithTextBox,
  docxWithTrackedChange,
  docxWithVerticalAlignment,
  docxWithoutExtras,
  docxWithReferences,
  docxWithSections,
} from '../../../e2e/fixtures.js'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  LineSpacingKind,
  TextAlignment,
  paragraphAttrsFrom,
  paragraphDraftFrom,
} from '@services/document/paragraph-format.js'
import { effectiveAttrs } from '@services/document/style-cascade.js'
import type { StyleSheet } from '@services/document/styles.js'
import { SidecarClient } from './client.js'
import { SidecarMethod } from './protocol.js'
import { sidecarPathIn } from './locate.js'

const executable = sidecarPathIn(process.cwd())
const published = await access(executable, constants.X_OK).then(
  () => true,
  () => false,
)

const client = new SidecarClient(() => Promise.resolve(executable))
afterAll(() => client.dispose())

const editorExtensions = '../../renderer/document/editor-extensions.js'
const { buildEditorExtensions } = (await import(editorExtensions)) as {
  buildEditorExtensions: (onSearchStatusChange: (status: unknown) => void) => Extensions
}

const schema = getSchema(buildEditorExtensions(() => {}))

const inflate = promisify(inflateRaw)

/** From the ZIP local headers, without a library. */
async function documentXmlOf(zip: Uint8Array): Promise<string> {
  const bytes = Buffer.from(zip)

  for (let i = 0; i + 30 <= bytes.length; i++) {
    if (bytes.readUInt32LE(i) !== 0x04034b50) continue

    const method = bytes.readUInt16LE(i + 8)
    const compressed = bytes.readUInt32LE(i + 18)
    const nameLength = bytes.readUInt16LE(i + 26)
    const extraLength = bytes.readUInt16LE(i + 28)
    const name = bytes.subarray(i + 30, i + 30 + nameLength).toString('utf8')
    if (name !== 'word/document.xml') continue

    const start = i + 30 + nameLength + extraLength
    const data = bytes.subarray(start, start + compressed)
    return method === 0 ? data.toString('utf8') : (await inflate(data)).toString('utf8')
  }

  throw new Error('o arquivo gravado não tem word/document.xml')
}

/** As `getJSON()` returns it. */
function throughEditor(doc: unknown): unknown {
  return ProseMirrorNode.fromJSON(schema, doc).toJSON()
}

/**
 * The model reduced to what the fingerprint compares. A hand-written mirror of `Node.Fingerprint()`
 * (`Nodes.cs`), so a bug in the normalizer does not cancel out on both sides. An attribute the
 * schema does not know disappears on one side only, and the comparison catches it.
 */
function asFingerprinted(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(asFingerprinted)
  if (value === null || typeof value !== 'object') return value

  // A note reference is identified by what it points to (`kind`, `nid`, `mark`); NotesWriter
  // compares the body.
  const record = value as Record<string, unknown>
  if (record['type'] === 'noteRef') {
    const attrs = (record['attrs'] ?? {}) as Record<string, unknown>
    const rest = Object.fromEntries(Object.entries(record).filter(([key]) => key !== 'content'))
    return fingerprintEntries({
      ...rest,
      attrs: { kind: attrs['kind'], nid: attrs['nid'], mark: attrs['mark'] },
    })
  }

  // An equation is identified by its OMML: MathML, LaTeX and list come from it in the sidecar.
  if (record['type'] === 'math') {
    const attrs = (record['attrs'] ?? {}) as Record<string, unknown>
    return fingerprintEntries({ ...record, attrs: { omml: attrs['omml'] } })
  }

  return fingerprintEntries(record)
}

function fingerprintEntries(value: Record<string, unknown>): Record<string, unknown> {

  const node: Record<string, unknown> = {}

  for (const [key, entry] of Object.entries(value)) {
    if (key === 'attrs') {
      const attrs = Object.entries((entry ?? {}) as Record<string, unknown>).filter(
        ([name, item]) => name !== 'oid' && name !== 'sectionBreak' && item !== null && item !== undefined,
      )
      if (attrs.length > 0) node['attrs'] = Object.fromEntries(attrs.map(([name, item]) => [name, item]))
      continue
    }

    if (key === 'marks' && Array.isArray(entry)) {
      node['marks'] = entry
        .map(asFingerprinted)
        .sort((left, right) => typeOf(left).localeCompare(typeOf(right)))
      continue
    }

    if (key === 'content' && Array.isArray(entry)) {
      node['content'] = mergeNeighbouringText(entry.map(asFingerprinted))
      continue
    }

    node[key] = asFingerprinted(entry)
  }

  return node
}

function typeOf(node: unknown): string {
  const type = (node as { type?: unknown } | null)?.type
  return typeof type === 'string' ? type : ''
}

/** Neighbouring text nodes with the same marks are a single text. */
function mergeNeighbouringText(content: unknown[]): unknown[] {
  const merged: Array<Record<string, unknown>> = []

  for (const child of content) {
    const node = child as Record<string, unknown>
    const previous = merged.at(-1)

    if (
      previous !== undefined &&
      node['type'] === 'text' &&
      previous['type'] === 'text' &&
      stable(previous['marks']) === stable(node['marks'])
    ) {
      previous['text'] = `${String(previous['text'] ?? '')}${String(node['text'] ?? '')}`
      continue
    }

    merged.push({ ...node })
  }

  return merged
}

/** Stable key order, to compare marks as text. */
function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`
  if (value === null || typeof value !== 'object') return JSON.stringify(value) ?? 'null'

  const entries = Object.entries(value as Record<string, unknown>).sort(([left], [right]) =>
    left.localeCompare(right),
  )
  return `{${entries.map(([key, entry]) => `${key}:${stable(entry)}`).join(',')}}`
}

interface OpenReply {
  readonly model: {
    readonly page: unknown
    readonly sections?: unknown
    readonly doc: unknown
    readonly styles: StyleSheet
  }
}

interface SaveReply {
  readonly preservedBlocks: number
  readonly rewrittenBlocks: number
}

async function openSaveAndCount(bytes: Buffer): Promise<SaveReply & { readonly bytes: Uint8Array }> {
  const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const { model } = opened.result as OpenReply

  const saved = await client.request(
    SidecarMethod.DocxSave,
    {
      page: model.page,
      ...(model.sections === undefined ? {} : { sections: model.sections }),
      doc: throughEditor(model.doc),
    },
    new Uint8Array(bytes),
  )

  return { ...(saved.result as SaveReply), bytes: saved.binary }
}

/** In order. The SDK closes the empty element with `" />"`, the only difference in writing. */
function sectionsOf(xml: string): string[] {
  return (xml.match(/<w:sectPr[ >][\s\S]*?<\/w:sectPr>/g) ?? []).map((section) =>
    section.replaceAll(' />', '/>'),
  )
}

describe.skipIf(!published)('impressão digital entre o editor e o sidecar', () => {
  // One case per corpus structure, because each diverges in its own way.
  const documents: Array<[string, () => Promise<Buffer>]> = [
    ['parágrafos com comentário ancorado', docxWithComment],
    // The anchor becomes a node and the reply's does not, on both sides.
    ['conversa de comentários com resposta e resolvido', docxWithCommentThread],
    ['parágrafo sem nada em volta', docxWithoutExtras],
    ['imagem esticada no fluxo do texto', docxWithStretchedImage],
    ['caixa de texto', docxWithTextBox],
    ['espaçamento dos dois lados', docxWithSpacingOnBothSides],
    ['cabeçalho em grade', docxWithHeaderGrid],
    ['lista com marcador', docxWithBulletList],
    // Levels, composite text, continuation and restart.
    ['lista multinível com reinício', docxWithMultilevelList],
    ['tabela com tabela aninhada', docxWithTable],
    ['sobrescrito e subscrito', docxWithVerticalAlignment],
    // Column width, shading, border, horizontal merge and header row in the table; alt text on the
    // image.
    ['tabela com sombreamento, borda e cabeçalho', docxWithStyledCells],
    ['imagem com texto alternativo', docxWithDescribedImage],
    // The block carries only direct formatting, and the rest belongs to the styles.
    ['estilos nomeados', () => docxWithNamedStyles()],
    ['formatação direta por cima dos estilos', docxWithDirectOverStyles],
    // Table of contents, bookmarks (including hidden ones), fields and internal link.
    ['referências do Word', () => docxWithReferences()],
    // Three sections, with an empty mark, a mark on a paragraph with text, and inheritance.
    ['seções', docxWithSections],
    // Revised insertion, deletion, paragraph mark, move and row.
    ['controle de alterações', () => docxWithTrackedChange()],
    // The note reference carries the body inside, and the fingerprint does not see it.
    ['nota de rodapé', () => docxWithFootnote()],
    // The equation carries its OMML as identity, and the derived MathML does not count.
    ['equações em linha e de exibição', docxWithEquations],
  ]

  it.each(documents)('abrir e salvar %s não reescreve bloco nenhum', async (_name, build) => {
    const bytes = await build()
    const result = await openSaveAndCount(bytes)

    expect(result.rewrittenBlocks).toBe(0)
    expect(result.preservedBlocks).toBeGreaterThan(0)
  })

  it('abrir e salvar um documento de seções devolve cada w:sectPr byte a byte', async () => {
    // Opened and saved without editing, no `w:sectPr` changes, not even the body's, which is the
    // last section.
    const bytes = await docxWithSections()
    const result = await openSaveAndCount(bytes)

    expect(result.rewrittenBlocks).toBe(0)
    expect(sectionsOf(await documentXmlOf(result.bytes))).toEqual(sectionsOf(await documentXmlOf(bytes)))
  })

  it.each(documents)('o modelo de %s volta do schema como o sidecar o leu', async (_name, build) => {
    // The counter says it passed; comparing the trees says why.
    const bytes = await build()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    expect(asFingerprinted(throughEditor(model.doc))).toEqual(asFingerprinted(model.doc))
  })

  it('o corpo da nota volta do schema como o sidecar o leu', async () => {
    // NotesWriter compares the note body block by block: an attribute lost in the schema would
    // rewrite the note on every save.
    const bytes = await docxWithFootnote()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const bodies = (doc: unknown): unknown[] => {
      const found: unknown[] = []
      const walk = (node: unknown): void => {
        const record = node as { type?: string; content?: unknown[] }
        if (record.type === 'noteRef') found.push(record.content)
        for (const child of record.content ?? []) walk(child)
      }
      walk(doc)
      return found
    }

    const read = bodies(model.doc)
    expect(read).toHaveLength(1)
    expect(asFingerprinted(bodies(throughEditor(model.doc)))).toEqual(asFingerprinted(read))
  })

  it('a saída do diálogo de parágrafo é o que o arquivo recebe', async () => {
    // The only path where the editor invents paragraph attributes: `lineHeight` as a CSS measure,
    // `textAlign` and `indent: 0`. The dialog factor does not go raw into the attribute, or the
    // writer would divide it by the font's natural height.
    const bytes = await docxWithoutExtras()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const doc = model.doc as { content: Array<{ type: string; attrs: Record<string, unknown> }> }
    const first = doc.content[0]!
    const attrs = paragraphAttrsFrom(
      {
        ...DEFAULT_PARAGRAPH_DRAFT,
        align: TextAlignment.Justify,
        lineSpacingKind: LineSpacingKind.Multiple,
        lineSpacingValue: 1.5,
        spaceAfter: 6,
      },
      first.attrs,
    )
    const edited = {
      ...(model.doc as Record<string, unknown>),
      content: [{ ...first, attrs: { ...first.attrs, ...attrs } }, ...doc.content.slice(1)],
    }

    const saved = await client.request(
      SidecarMethod.DocxSave,
      { page: model.page, doc: throughEditor(edited) },
      new Uint8Array(bytes),
    )
    const result = saved.result as SaveReply & { inventory: { lost: string[] } }

    // One rewritten block and no loss: out-of-range line spacing would go to the inventory.
    expect(result.rewrittenBlocks).toBe(1)
    expect(result.inventory.lost).toEqual([])

    // `w:line` in 240ths with `w:lineRule="auto"`: 1.5 lines is 360.
    const xml = await documentXmlOf(saved.binary)
    expect(xml).toContain('w:line="360"')

    // And reopening gives the dialog back what it chose.
    const reopened = await client.request(SidecarMethod.DocxOpen, {}, saved.binary)
    const back = (reopened.result as OpenReply).model.doc as {
      content: Array<{ attrs: Record<string, unknown> }>
    }
    const draft = paragraphDraftFrom(back.content[0]!.attrs)

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Multiple)
    expect(draft.lineSpacingValue).toBe(1.5)
    expect(draft.align).toBe(TextAlignment.Justify)
    expect(draft.spaceAfter).toBe(6)
  })

  it('o diálogo de parágrafo aplicado sem mudança não reescreve bloco nenhum', async () => {
    // "OK" without changes does not turn what the style inherits into direct formatting.
    const bytes = await docxWithDirectOverStyles()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const doc = model.doc as { content: Array<{ type: string; attrs?: Record<string, unknown> }> }
    const content = doc.content.map((block) => {
      const attrs = block.attrs ?? {}
      const effective = effectiveAttrs(block, model.styles)
      return { ...block, attrs: { ...attrs, ...paragraphAttrsFrom(paragraphDraftFrom(effective), attrs, effective) } }
    })

    const saved = await client.request(
      SidecarMethod.DocxSave,
      { page: model.page, doc: throughEditor({ ...(model.doc as Record<string, unknown>), content }) },
      new Uint8Array(bytes),
    )
    const result = saved.result as SaveReply

    expect(result.rewrittenBlocks).toBe(0)
    expect(result.preservedBlocks).toBe(doc.content.length)
  })

  it('cada bloco do modelo chega ao editor com identidade', async () => {
    // The `oid` links the screen block to the file's `w:p`.
    const bytes = await docxWithComment()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const blocks = (model.doc as { content: Array<{ attrs?: Record<string, unknown> }> }).content
    expect(blocks.length).toBeGreaterThan(0)
    expect(blocks.every((block) => typeof block.attrs?.['oid'] === 'string')).toBe(true)
  })
})
