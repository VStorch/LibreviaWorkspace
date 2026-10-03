import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'
import { numberLists, LIST_TYPES, type ListInfo, type ListTreeReader } from './list-numbering.js'
import type { DocumentModel, DocumentNode } from './model.js'
import { NoteKind, noteCounter } from './notes.js'

/** Do **modelo**, e não do HTML do editor, que traz decorações e atributos que só o editor entende. */

export interface ExportNote {
  readonly kind: NoteKind
  readonly label: string
  /** Âncora da nota (`nota-rodape-1`, `nota-fim-1`) — a da referência leva `ref-` antes. */
  readonly id: string
  readonly body: readonly DocumentNode[]
}

export interface ExportListItem {
  readonly label: string
  readonly value: number
}

export interface ExportSource {
  readonly doc: DocumentNode
  /** As notas, na ordem do texto; as de rodapé antes das de fim no rodapé da exportação. */
  readonly notes: readonly ExportNote[]
  readonly noteOf: ReadonlyMap<DocumentNode, ExportNote>
  readonly listOf: ReadonlyMap<DocumentNode, ListInfo>
  readonly itemOf: ReadonlyMap<DocumentNode, ExportListItem>
  /** Os marcadores que algum link interno aponta (`#nome`). */
  readonly linkTargets: ReadonlySet<string>
}

const JSON_READER: ListTreeReader<DocumentNode> = {
  typeOf: (node) => node.type,
  attrsOf: (node) => node.attrs ?? {},
  childrenOf: (node) => node.content ?? [],
}

type Mark = NonNullable<DocumentNode['marks']>[number]

const hasMark = (node: DocumentNode, type: string): boolean =>
  node.marks?.some((mark) => mark.type === type) ?? false

function blockRevisionKind(value: unknown): string | null {
  if (typeof value !== 'object' || value === null) return null
  const kind = (value as { kind?: unknown }).kind
  return kind === 'ins' || kind === 'del' ? kind : null
}

const NEEDS_BLOCK = new Set(['tableCell', 'tableHeader', 'listItem', 'noteRef', 'blockquote'])

const MERGEABLE = new Set(['paragraph', 'heading'])

/** Nós que não vão a nenhuma exportação: as pontas dos comentários. */
const DROPPED = new Set(['commentStart', 'commentEnd'])

/**
 * O que "Aceitar todas" faria, sem passar pelo editor. Os comentários saem,
 * menos no ODT (`keepComments`), que os leva como anotações.
 */
export function finalDocument(node: DocumentNode, keepComments = false): DocumentNode {
  let content = node.content === undefined ? undefined : finalChildren(node.content, keepComments)
  // Um bloco que o esquema não deixa vazio (a célula, o item, a nota) volta com
  // um parágrafo quando tudo o que tinha era excluído.
  if (content?.length === 0 && NEEDS_BLOCK.has(node.type)) content = [{ type: 'paragraph' }]
  const marks = node.marks?.filter((mark) => mark.type !== 'insertion')
  const { content: _content, marks: _marks, ...rest } = node
  void _content
  void _marks
  return {
    ...rest,
    ...(content === undefined ? {} : { content }),
    ...(marks === undefined || marks.length === 0 ? {} : { marks }),
  }
}

function finalChildren(children: readonly DocumentNode[], keepComments: boolean): DocumentNode[] {
  const kept: DocumentNode[] = []
  for (const child of children) {
    if ((!keepComments && DROPPED.has(child.type)) || hasMark(child, 'deletion')) continue
    if (child.type === 'tableRow' && blockRevisionKind(child.attrs?.['rowRevision']) === 'del') continue
    kept.push(finalDocument(child, keepComments))
  }

  // A marca de parágrafo excluída: o texto dele continua no parágrafo seguinte.
  const merged: DocumentNode[] = []
  let pending: DocumentNode | null = null
  for (const child of kept) {
    if (pending !== null && MERGEABLE.has(child.type)) {
      merged.push({ ...child, content: [...(pending.content ?? []), ...(child.content ?? [])] })
      pending = null
    } else {
      if (pending !== null) merged.push(pending)
      pending = null
      if (MERGEABLE.has(child.type) && blockRevisionKind(child.attrs?.['markRevision']) === 'del') {
        pending = child
        continue
      }
      merged.push(child)
    }
  }
  if (pending !== null) merged.push(pending)
  return merged
}

/** Em pré-ordem, a ordem do texto. */
export function walk(node: DocumentNode, visit: (node: DocumentNode) => void): void {
  visit(node)
  for (const child of node.content ?? []) walk(child, visit)
}

export function prepareExport(
  model: Pick<DocumentModel, 'doc' | 'notes'>,
  options: { readonly keepComments?: boolean } = {},
): ExportSource {
  const doc = finalDocument(model.doc, options.keepComments === true)

  const noteOf = new Map<DocumentNode, ExportNote>()
  const notes: ExportNote[] = []
  const label = noteCounter(model.notes)
  const ordinals = { [NoteKind.Footnote]: 0, [NoteKind.Endnote]: 0 }
  let section = 0
  const listNodes: DocumentNode[] = []
  const itemNodes: DocumentNode[] = []
  const linkTargets = new Set<string>()

  walk(doc, (node) => {
    if (node.type === 'noteRef') {
      const kind = node.attrs?.['kind'] === NoteKind.Endnote ? NoteKind.Endnote : NoteKind.Footnote
      const mark = typeof node.attrs?.['mark'] === 'string' ? node.attrs['mark'] : null
      ordinals[kind] += 1
      const note: ExportNote = {
        kind,
        label: label({ kind, mark, section }),
        id: `${kind === NoteKind.Endnote ? 'nota-fim' : 'nota-rodape'}-${ordinals[kind]}`,
        body: node.content ?? [],
      }
      noteOf.set(node, note)
      notes.push(note)
    }
    if (typeof node.attrs?.['sectionBreak'] === 'string') section += 1
    if (LIST_TYPES.includes(node.type)) listNodes.push(node)
    if (node.type === 'listItem') itemNodes.push(node)
    for (const mark of node.marks ?? []) {
      const href = linkHref(mark)
      if (href?.startsWith('#') === true) linkTargets.add(href.slice(1))
    }
  })

  // `numberLists` conta na mesma pré-ordem.
  const numbering = numberLists(doc, JSON_READER)
  const listOf = new Map<DocumentNode, ListInfo>()
  listNodes.forEach((node, index) => {
    const info = numbering.lists[index]
    if (info !== undefined) listOf.set(node, info)
  })
  const itemOf = new Map<DocumentNode, ExportListItem>()
  itemNodes.forEach((node, index) => {
    itemOf.set(node, { label: numbering.labels[index] ?? '', value: numbering.values[index] ?? index + 1 })
  })

  const ordered = [
    ...notes.filter((note) => note.kind === NoteKind.Footnote),
    ...notes.filter((note) => note.kind === NoteKind.Endnote),
  ]
  return { doc, notes: ordered, noteOf, listOf, itemOf, linkTargets }
}

function linkHref(mark: Mark): string | null {
  if (mark.type !== 'link') return null
  const href = mark.attrs?.['href']
  return typeof href === 'string' ? href : null
}

/** Âncora interna ou um protocolo que o editor abre; `javascript:`, `file:` e `data:` viram texto. */
export function safeHref(mark: Mark): string | null {
  const href = linkHref(mark)?.trim()
  if (href === undefined || href === null || href === '') return null
  if (href.startsWith('#')) return href.length > 1 ? href : null
  try {
    const parsed = new URL(href)
    return (ALLOWED_EXTERNAL_PROTOCOLS as readonly string[]).includes(parsed.protocol) ? href : null
  } catch {
    return null
  }
}

/** A imagem embutida, dividida em tipo e dados; `null` se não for imagem em `data:`. */
export function imageData(src: unknown): { readonly mime: string; readonly base64: string } | null {
  if (typeof src !== 'string') return null
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/=\s]+)$/i.exec(src)
  if (match === null) return null
  return { mime: match[1]!.toLowerCase(), base64: match[2]!.replace(/\s+/g, '') }
}

const EXTENSIONS: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/bmp': 'bmp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tif',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'image/emf': 'emf',
  'image/wmf': 'wmf',
}

export function imageExtension(mime: string): string {
  return EXTENSIONS[mime] ?? 'bin'
}

/** O nível de um parágrafo do sumário, pelo estilo (`TOC2`, `toc 2`, `Sumário2`). */
export function tocLevelOf(node: DocumentNode): number {
  const style = node.attrs?.['styleId']
  const match = typeof style === 'string' ? /(\d)\s*$/.exec(style) : null
  return match === null ? 1 : Math.max(1, Number(match[1]))
}

/** Sem o número de página: na página web ele apontaria para lugar nenhum. */
export function withoutPageNumbers(content: readonly DocumentNode[]): DocumentNode[] {
  const kept = content.filter(
    (node) => !(node.type === 'field' && /^\s*PAGEREF\b/i.test(String(node.attrs?.['instr'] ?? ''))),
  )
  // O tab que separava o título do número de página fica sobrando no fim.
  while (kept.length > 0) {
    const last = kept[kept.length - 1]!
    if (last.type !== 'text' || (last.text ?? '').trimEnd() !== '') break
    kept.pop()
  }
  const last = kept[kept.length - 1]
  if (last?.type === 'text' && last.text !== undefined) {
    kept[kept.length - 1] = { ...last, text: last.text.replace(/\t\s*[0-9ivxlcdm]*\s*$/i, '').trimEnd() }
  }
  return kept
}

export function exportTitle(model: Pick<DocumentModel, 'properties'>, fileName: string): string {
  const title = model.properties?.title?.trim()
  if (title !== undefined && title !== '') return title
  const dot = fileName.lastIndexOf('.')
  return dot > 0 ? fileName.slice(0, dot) : fileName
}

/** O parágrafo que só carrega a marca de uma seção não tem texto a exportar. */
export function isSectionMarkOnly(node: DocumentNode): boolean {
  return node.attrs?.['sectionMark'] === true && (node.content ?? []).length === 0
}
