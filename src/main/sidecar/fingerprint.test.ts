/**
 * O contrato da impressão digital, com os dois lados de verdade: o documento é lido
 * pelo sidecar publicado, atravessa o schema montado com as extensões reais do
 * editor e volta ao sidecar. O schema reescreve o modelo (materializa atributos,
 * funde textos, ordena marcas, descarta o desconhecido), e uma diferença que só
 * aparece nessa costura regenera o documento sem falhar em lugar nenhum.
 *
 * Dois critérios: zero blocos reescritos, e as duas árvores iguais no que a
 * impressão digital compara, porque só o número deixaria passar o que o schema e a
 * impressão digital ignorassem juntos.
 *
 * Sem o `Editor` do Tiptap, que precisa de DOM: o que ele faz com `content` é
 * `Node.fromJSON` sobre o schema, e `getJSON()` é o `toJSON()`. As extensões vêm
 * por caminho em variável porque `tsconfig.node.json` não conhece `src/renderer`.
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

/** O `word/document.xml` do `.docx` gravado, pelos cabeçalhos locais do ZIP, sem biblioteca. */
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

/** O modelo depois de atravessar o schema do editor, como `getJSON()` o devolve. */
function throughEditor(doc: unknown): unknown {
  return ProseMirrorNode.fromJSON(schema, doc).toJSON()
}

/**
 * O modelo reduzido ao que a impressão digital compara. Espelho de
 * `Node.Fingerprint()` (`Nodes.cs`) escrito à mão, para que um erro no normalizador
 * não se cancele nos dois lados. Um atributo que o schema não conhece some de um
 * lado só, e a comparação acusa.
 */
function asFingerprinted(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(asFingerprinted)
  if (value === null || typeof value !== 'object') return value

  // A referência de nota vale pelo que aponta (`kind`, `nid`, `mark`); o corpo NotesWriter compara.
  const record = value as Record<string, unknown>
  if (record['type'] === 'noteRef') {
    const attrs = (record['attrs'] ?? {}) as Record<string, unknown>
    const rest = Object.fromEntries(Object.entries(record).filter(([key]) => key !== 'content'))
    return fingerprintEntries({
      ...rest,
      attrs: { kind: attrs['kind'], nid: attrs['nid'], mark: attrs['mark'] },
    })
  }

  // A equação vale pelo OMML: MathML, LaTeX e lista saem dele no sidecar.
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

/** Nós de texto vizinhos com as mesmas marcas são um texto só. */
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

/** JSON com as chaves em ordem estável, para comparar marcas como texto. */
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

/**
 * Os `w:sectPr` do documento, em ordem. O SDK fecha o elemento vazio com `" />"`, a
 * única diferença de escrita.
 */
function sectionsOf(xml: string): string[] {
  return (xml.match(/<w:sectPr[ >][\s\S]*?<\/w:sectPr>/g) ?? []).map((section) =>
    section.replaceAll(' />', '/>'),
  )
}

describe.skipIf(!published)('impressão digital entre o editor e o sidecar', () => {
  // Um caso por estrutura do corpus, porque cada uma diverge de um jeito.
  const documents: Array<[string, () => Promise<Buffer>]> = [
    ['parágrafos com comentário ancorado', docxWithComment],
    // A âncora vira nó, e a da resposta não — nos dois lados.
    ['conversa de comentários com resposta e resolvido', docxWithCommentThread],
    ['parágrafo sem nada em volta', docxWithoutExtras],
    ['imagem esticada no fluxo do texto', docxWithStretchedImage],
    ['caixa de texto', docxWithTextBox],
    ['espaçamento dos dois lados', docxWithSpacingOnBothSides],
    ['cabeçalho em grade', docxWithHeaderGrid],
    ['lista com marcador', docxWithBulletList],
    // Níveis, texto composto, continuação e reinício.
    ['lista multinível com reinício', docxWithMultilevelList],
    ['tabela com tabela aninhada', docxWithTable],
    ['sobrescrito e subscrito', docxWithVerticalAlignment],
    // Largura de coluna, sombreamento, borda, mesclagem horizontal e linha de
    // cabeçalho na tabela; texto alternativo na imagem.
    ['tabela com sombreamento, borda e cabeçalho', docxWithStyledCells],
    ['imagem com texto alternativo', docxWithDescribedImage],
    // O bloco leva só a formatação direta, e o resto é dos estilos.
    ['estilos nomeados', () => docxWithNamedStyles()],
    ['formatação direta por cima dos estilos', docxWithDirectOverStyles],
    // Sumário, marcadores (inclusive os ocultos), campos e link interno.
    ['referências do Word', () => docxWithReferences()],
    // Três seções, com marca vazia, marca em parágrafo com texto e herança.
    ['seções', docxWithSections],
    // Inserção, exclusão, marca de parágrafo, movimentação e linha revisadas.
    ['controle de alterações', () => docxWithTrackedChange()],
    // A referência de nota leva o corpo dentro, e a impressão digital não o vê.
    ['nota de rodapé', () => docxWithFootnote()],
    // A equação leva o OMML como identidade, e o MathML derivado não conta.
    ['equações em linha e de exibição', docxWithEquations],
  ]

  it.each(documents)('abrir e salvar %s não reescreve bloco nenhum', async (_name, build) => {
    const bytes = await build()
    const result = await openSaveAndCount(bytes)

    expect(result.rewrittenBlocks).toBe(0)
    expect(result.preservedBlocks).toBeGreaterThan(0)
  })

  it('abrir e salvar um documento de seções devolve cada w:sectPr byte a byte', async () => {
    // Aberto e gravado sem editar, nenhum `w:sectPr` muda, nem o do corpo, que é a última seção.
    const bytes = await docxWithSections()
    const result = await openSaveAndCount(bytes)

    expect(result.rewrittenBlocks).toBe(0)
    expect(sectionsOf(await documentXmlOf(result.bytes))).toEqual(sectionsOf(await documentXmlOf(bytes)))
  })

  it.each(documents)('o modelo de %s volta do schema como o sidecar o leu', async (_name, build) => {
    // O contador diz que passou; a comparação das árvores diz por quê.
    const bytes = await build()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    expect(asFingerprinted(throughEditor(model.doc))).toEqual(asFingerprinted(model.doc))
  })

  it('o corpo da nota volta do schema como o sidecar o leu', async () => {
    // O corpo da nota NotesWriter compara bloco a bloco: um atributo perdido no schema
    // reescreveria a nota a cada gravação.
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
    // O único caminho em que o editor inventa atributos de parágrafo: `lineHeight` como
    // medida de CSS, `textAlign` e `indent: 0`. O fator do diálogo não vai cru ao
    // atributo, senão o gravador o divide pela altura natural da fonte.
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

    // Um bloco reescrito e nenhuma perda: entrelinha fora da faixa iria ao inventário.
    expect(result.rewrittenBlocks).toBe(1)
    expect(result.inventory.lost).toEqual([])

    // `w:line` em 240-avos com `w:lineRule="auto"`: 1,5 linha são 360.
    const xml = await documentXmlOf(saved.binary)
    expect(xml).toContain('w:line="360"')

    // E reabrir devolve ao diálogo o que ele escolheu.
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
    // "OK" sem mexer não transforma o herdado do estilo em direto.
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
    // O `oid` liga o bloco da tela ao `w:p` do arquivo.
    const bytes = await docxWithComment()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const blocks = (model.doc as { content: Array<{ attrs?: Record<string, unknown> }> }).content
    expect(blocks.length).toBeGreaterThan(0)
    expect(blocks.every((block) => typeof block.attrs?.['oid'] === 'string')).toBe(true)
  })
})
