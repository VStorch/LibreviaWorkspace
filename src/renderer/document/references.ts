import type { Editor } from '@tiptap/react'
import type { Node as ProseMirrorNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { pageLabel } from '@services/document/band.js'
import { hiddenBookmarkName, nextBookmarkId } from '@services/document/bookmarks.js'
import {
  fieldArgument,
  fieldKind,
  fieldSwitch,
  sequenceNumbers,
  tocLevels,
  tocLinks,
  tocOmitsPages,
} from '@services/document/fields.js'
import type { PageSetup } from '@services/document/model.js'
import type { NoteKind } from '@services/document/notes.js'
import { outlineOf } from '@services/document/outline.js'
import {
  ensureCaptionStyle,
  ensureTocHeadingStyle,
  ensureTocStyle,
} from '@services/document/reference-styles.js'
import type { StyleSheet } from '@services/document/styles.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { bookmarksOf } from './extensions/bookmark.js'
import { noteLabelsOf, noteRefsOf, textBetweenWithoutNotes } from './extensions/note-ref.js'
import { DEFAULT_TOC_INSTRUCTION } from './extensions/table-of-contents.js'
import { readPendingSelection, textStartOf } from './extensions/zero-width.js'
import { outlineBlocksOf } from './outline-blocks.js'
import { drawnSheet, type PageLayout, type PageStart } from './usePagination.js'

/** Lido na hora do comando: a paginação muda a cada linha, e um número de página velho é o erro que o sumário não pode ter. */
export interface ReferenceContext {
  readonly layout: PageLayout
  readonly page: PageSetup
  /** O número da folha sai no formato da seção dela. Ausente, vale `page`. */
  readonly sections?: readonly PageSetup[]
  readonly styles: StyleSheet
  readonly setStyles: (styles: StyleSheet) => void
  readonly t: (key: MessageKey) => string
  /** Marcadores que existem no arquivo fora dos nós — ver `DocumentModel.outsideBookmarks`. */
  readonly outsideBookmarks?: readonly string[]
}

/** Faz o ProseMirror ler agora a seleção que o navegador já mudou. */
export function flushSelection(editor: Editor): void {
  readPendingSelection(editor.view)
}

/** A posição do documento em que a folha começa, ou nulo se o layout envelheceu. */
function positionOfStart(doc: ProseMirrorNode, start: PageStart): number | null {
  if (start.blockIndex >= doc.childCount) return null

  let pos = 0
  for (let index = 0; index < start.blockIndex; index++) pos += doc.child(index).nodeSize
  const block = doc.child(start.blockIndex)

  if (start.offset !== undefined) return pos + 1 + start.offset
  if (start.childIndex !== undefined) {
    let inner = pos + 1
    for (let index = 0; index < start.childIndex && index < block.childCount; index++) {
      inner += block.child(index).nodeSize
    }
    return inner
  }
  return pos
}

/** Com o reinício e o formato da seção em que a folha cai, como o campo `PAGE` escreve. */
export function sheetLabel(context: ReferenceContext, sheet: number): string {
  const plan = context.layout.sheets[drawnSheet(context.layout, sheet - 1)]
  if (plan === undefined) return pageLabel(context.page, sheet)
  const section = context.sections?.[plan.section] ?? context.page
  return pageLabel({ ...section, pageNumberStart: plan.number }, 1)
}

/** A folha (de 1 em diante) em que a posição cai, pelos cortes da paginação. */
export function sheetAt(doc: ProseMirrorNode, starts: readonly PageStart[], pos: number): number {
  let sheet = 1
  for (const start of starts) {
    const at = positionOfStart(doc, start)
    if (at === null || at > pos) break
    sheet += 1
  }
  return sheet
}

/** O texto entre duas posições, com o resultado dos campos no lugar deles. */
function textBetween(doc: ProseMirrorNode, from: number, to: number): string {
  return textBetweenWithoutNotes(doc, from, to, ' ', (leaf) =>
    leaf.type.name === 'field' ? String(leaf.attrs['result'] ?? '') : '',
  )
}

/** O que o `NOTEREF` mostra; `null` sem nota no trecho, e o campo fica como está. */
export function noteNumberIn(
  doc: ProseMirrorNode,
  labels: readonly string[],
  from: number,
  to: number,
): string | null {
  const index = noteRefsOf(doc).findIndex(({ pos }) => pos >= from && pos < to)
  return index < 0 ? null : (labels[index] ?? null)
}

/** Os tipos que dependem de onde o texto cai na folha. */
const PAGE_KINDS = new Set(['PAGE', 'PAGEREF', 'NUMPAGES'])

/** Os tipos que o editor sabe recalcular. O resto fica como o Word o deixou. */
const UPDATABLE = new Set(['PAGE', 'PAGEREF', 'NUMPAGES', 'REF', 'SEQ', 'NOTEREF'])

export interface FieldUpdate {
  readonly changed: number
  /** Algum campo de página foi recalculado — e a paginação pode mudar com ele. */
  readonly pageDependent: boolean
}

/**
 * O F9 do Word, na ordem dele: `SEQ` primeiro, porque o `REF` a uma legenda cita
 * o número; depois as referências; e as páginas por último. O `SEQ` conta o
 * documento inteiro, mas só os de dentro do trecho mudam. Marcador que não
 * existe mais dá "Erro! Indicador não definido.", como no Word.
 */
export function updateFieldsIn(
  editor: Editor,
  context: ReferenceContext,
  from: number,
  to: number,
  kinds: ReadonlySet<string> = UPDATABLE,
): FieldUpdate {
  // O cursor recém-movido pode estar só no DOM: sem isto a transação o puxaria de volta.
  flushSelection(editor)
  const { state } = editor
  const doc = state.doc

  const fields: Array<{ pos: number; node: ProseMirrorNode; kind: string }> = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'field') fields.push({ pos, node, kind: fieldKind(String(node.attrs['instr'])) })
    return true
  })

  const sequences = fields.filter((field) => field.kind === 'SEQ')
  const numbers = sequenceNumbers(sequences.map((field) => String(field.node.attrs['instr'])))
  const sequenceResult = new Map(sequences.map((field, index) => [field.pos, numbers[index]!]))

  // A referência lê o marcador **depois** das sequências: "Figura 1" que virou "Figura 2".
  const updatedSequenceDoc = (() => {
    const tr = state.tr
    for (const [pos, result] of sequenceResult) tr.setNodeAttribute(pos, 'result', result)
    return tr.doc
  })()

  const bookmarks = new Map(bookmarksOf(updatedSequenceDoc).map((bookmark) => [bookmark.name, bookmark]))
  const missing = context.t('references.field.missingBookmark')
  // O marcador fora dos nós não está perdido: fica o resultado que o Word calculou.
  const outside = new Set(context.outsideBookmarks ?? [])
  const sheets = context.layout.pages
  const noteLabels = noteLabelsOf(state)

  const tr = state.tr
  let changed = 0
  let pageDependent = false

  for (const field of fields) {
    if (field.pos < from || field.pos >= to) continue
    if (!kinds.has(field.kind)) continue

    const instr = String(field.node.attrs['instr'])
    let result: string | null = null

    switch (field.kind) {
      case 'SEQ':
        result = fieldSwitch(instr, 'h') !== null ? '' : (sequenceResult.get(field.pos) ?? null)
        break
      case 'REF': {
        const name = fieldArgument(instr) ?? ''
        const target = bookmarks.get(name)
        if (target === undefined) {
          result = outside.has(name) ? null : missing
          break
        }
        const text = textBetween(updatedSequenceDoc, target.pos + 1, target.end ?? target.pos + 1)
        // `\# 0`: só o número do texto citado, a referência "só o número" do Word.
        result = fieldSwitch(instr, '#') === null ? text : (/(\d+)(?!.*\d)/.exec(text)?.[1] ?? text)
        break
      }
      case 'NOTEREF': {
        // O número da tela, com os reinícios por folha e por seção.
        const name = fieldArgument(instr) ?? ''
        const target = bookmarks.get(name)
        if (target === undefined) {
          result = outside.has(name) ? null : missing
          break
        }
        result = noteNumberIn(doc, noteLabels, target.pos, target.end ?? target.pos)
        break
      }
      case 'PAGEREF': {
        const name = fieldArgument(instr) ?? ''
        const target = bookmarks.get(name)
        if (target === undefined) {
          result = outside.has(name) ? null : missing
          break
        }
        result = sheetLabel(context, sheetAt(doc, context.layout.pageStarts, target.pos))
        pageDependent = true
        break
      }
      case 'PAGE':
        result = sheetLabel(context, sheetAt(doc, context.layout.pageStarts, field.pos))
        pageDependent = true
        break
      case 'NUMPAGES':
        result = String(sheets)
        pageDependent = true
        break
      default:
        break
    }

    if (result === null || result === field.node.attrs['result']) continue
    tr.setNodeAttribute(field.pos, 'result', result)
    changed += 1
  }

  if (changed > 0) editor.view.dispatch(tr)
  return { changed, pageDependent }
}

/**
 * Se algum campo depende da página, há um segundo passe quando a paginação
 * assenta: o texto novo pode empurrar uma linha para a folha seguinte. O Word
 * também faz os dois.
 */
export function updateFields(editor: Editor, context: ReferenceContext): FieldUpdate {
  const { from, to, empty } = editor.state.selection
  const range = empty ? { from: 0, to: editor.state.doc.content.size } : { from, to }
  const update = updateFieldsIn(editor, context, range.from, range.to)
  // Só quando algo mudou: senão o passe reescreveria campos na próxima digitação.
  if (update.pageDependent && update.changed > 0 && empty) arm(editor, 'all')
  return update
}

/** O F9 do documento corrige todos os campos de página; o sumário, só os dele. */
const pendingPagePass = new WeakMap<Editor, { scope: 'all' | 'toc'; doc: ProseMirrorNode }>()

function arm(editor: Editor, scope: 'all' | 'toc'): void {
  pendingPagePass.set(editor, { scope, doc: editor.state.doc })
}

/** Uma vez: o passe não pede outro. */
export function settlePageFields(editor: Editor, context: ReferenceContext): void {
  const pending = pendingPagePass.get(editor)
  if (pending === undefined) return
  pendingPagePass.delete(editor)
  // A pessoa voltou a escrever: o passe é de outro documento.
  if (pending.doc !== editor.state.doc) return
  const { scope } = pending

  if (scope === 'all') {
    updateFieldsIn(editor, context, 0, editor.state.doc.content.size, PAGE_KINDS)
    return
  }

  for (const { pos, node } of tablesOfContents(editor.state.doc)) {
    updateFieldsIn(editor, context, pos, pos + node.nodeSize, PAGE_KINDS)
  }
}

/** De trás para a frente, para uma inserção não deslocar as que faltam. */
function ensureBookmarks(tr: Transaction, positions: readonly number[], prefix: '_Toc' | '_Ref'): string[] {
  const schema: Schema = tr.doc.type.schema
  const names = new Array<string>(positions.length)
  const used = bookmarksOf(tr.doc).map((bookmark) => bookmark.name)
  const ids: string[] = []
  tr.doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd')
      ids.push(String(node.attrs['bid']))
    return true
  })

  const order = positions.map((pos, index) => ({ pos, index })).sort((left, right) => right.pos - left.pos)
  for (const { pos, index } of order) {
    const block = tr.doc.nodeAt(pos)
    if (block === null) continue

    let found: string | null = null
    block.forEach((child) => {
      const name = String(child.attrs['name'] ?? '')
      if (found === null && child.type.name === 'bookmarkStart' && name.startsWith(prefix)) found = name
    })
    if (found !== null) {
      names[index] = found
      continue
    }

    const name = hiddenBookmarkName(prefix, used)
    const bid = nextBookmarkId(ids)
    used.push(name)
    ids.push(bid)
    tr.insert(pos + block.nodeSize - 1, schema.nodes['bookmarkEnd']!.create({ bid }))
    tr.insert(pos + 1, schema.nodes['bookmarkStart']!.create({ name, bid }))
    names[index] = name
  }

  return names
}

function tablesOfContents(doc: ProseMirrorNode): Array<{ pos: number; node: ProseMirrorNode }> {
  const found: Array<{ pos: number; node: ProseMirrorNode }> = []
  doc.forEach((node, pos) => {
    if (node.type.name === 'tableOfContents') found.push({ pos, node })
  })
  return found
}

/** Como o Word: o texto do título, uma tabulação e o `PAGEREF`, dentro do link para o marcador `_Toc…`. */
function buildEntries(
  tr: Transaction,
  context: ReferenceContext,
  instr: string,
  sheet: StyleSheet,
): { entries: unknown[]; sheet: StyleSheet } {
  const { from, to } = tocLevels(instr)
  const links = tocLinks(instr)
  const omitPages = tocOmitsPages(instr)

  // Um sumário antigo com parágrafo em estilo de título se listaria a si mesmo.
  const inside = tablesOfContents(tr.doc).map(({ pos, node }) => [pos, pos + node.nodeSize] as const)
  const headings = outlineOf(outlineBlocksOf(tr.doc), sheet).filter(
    (heading) =>
      heading.level >= from &&
      heading.level <= to &&
      !inside.some(([start, end]) => heading.pos > start && heading.pos < end),
  )

  const names = ensureBookmarks(
    tr,
    headings.map((heading) => heading.pos),
    '_Toc',
  )

  let styles = sheet
  const entries: unknown[] = headings.map((heading, index) => {
    const ensured = ensureTocStyle(styles, heading.level)
    styles = ensured.sheet
    const name = names[index]!
    const marks = links ? [{ type: 'link', attrs: { href: `#${name}` } }] : []
    // A folha da paginação atual, que mediu o documento antes dos marcadores; o segundo passe corrige.
    const page = sheetLabel(context, sheetAt(tr.before, context.layout.pageStarts, heading.pos))
    return {
      type: 'paragraph',
      attrs: { styleId: ensured.id },
      content: [
        { type: 'text', text: omitPages ? heading.text : `${heading.text}\t`, marks },
        ...(omitPages
          ? []
          : [{ type: 'field', attrs: { instr: ` PAGEREF ${name} \\h `, result: page }, marks }]),
      ],
    }
  })

  if (entries.length === 0) {
    entries.push({
      type: 'paragraph',
      attrs: { styleId: null },
      content: [{ type: 'text', text: context.t('references.toc.empty') }],
    })
  }

  return { entries, sheet: styles }
}

/** Antes do bloco do cursor, ou no lugar dele se estiver vazio, com os títulos 1 a 3. */
export function insertTableOfContents(editor: Editor, context: ReferenceContext): void {
  const { state } = editor
  const tr = state.tr

  const heading = ensureTocHeadingStyle(context.styles)
  const { entries, sheet } = buildEntries(tr, context, DEFAULT_TOC_INSTRUCTION, heading.sheet)

  const title = {
    type: 'paragraph',
    attrs: { styleId: heading.id },
    content: [{ type: 'text', text: context.t('references.toc.title') }],
  }
  const node = state.schema.nodeFromJSON({
    type: 'tableOfContents',
    attrs: { instr: DEFAULT_TOC_INSTRUCTION, head: 1, sdt: true },
    content: [title, ...entries],
  })

  const $from = tr.doc.resolve(tr.mapping.map(state.selection.from))
  const top = $from.depth === 0 ? $from.pos : $from.before(1)
  const current = tr.doc.nodeAt(top)
  if (current !== null && current.isTextblock && current.content.size === 0) {
    tr.replaceWith(top, top + current.nodeSize, node)
  } else {
    tr.insert(top, node)
  }

  if (sheet !== context.styles) context.setStyles(sheet)
  editor.view.dispatch(tr.scrollIntoView())
  arm(editor, 'toc')
}

/**
 * O do cursor, ou o primeiro. O título, a instrução e o controle de conteúdo
 * ficam; as entradas são trocadas inteiras, como no "Atualizar sumário inteiro"
 * do Word. Falso quando não há sumário.
 */
export function updateTableOfContents(editor: Editor, context: ReferenceContext): boolean {
  const { state } = editor
  const all = tablesOfContents(state.doc)
  if (all.length === 0) return false

  const cursor = state.selection.from
  const chosen = all.find(({ pos, node }) => cursor > pos && cursor < pos + node.nodeSize) ?? all[0]!
  const tr = state.tr
  const instr = String(chosen.node.attrs['instr'] ?? DEFAULT_TOC_INSTRUCTION)
  const { entries, sheet } = buildEntries(tr, context, instr, context.styles)

  // Os marcadores novos podem ter entrado antes do sumário.
  const pos = tr.mapping.map(chosen.pos)
  const toc = tr.doc.nodeAt(pos)
  if (toc === null) return false

  const head = Math.min(Number(toc.attrs['head'] ?? 0), toc.childCount - 1)
  const kept: ProseMirrorNode[] = []
  for (let index = 0; index < head; index++) kept.push(toc.child(index))
  const fresh = entries.map((entry) => state.schema.nodeFromJSON(entry))

  tr.replaceWith(pos + 1, pos + toc.nodeSize - 1, [...kept, ...fresh])

  if (sheet !== context.styles) context.setStyles(sheet)
  editor.view.dispatch(tr)
  arm(editor, 'toc')
  return true
}

function sequencesIn(block: ProseMirrorNode, pos: number): Array<{ pos: number; label: string }> {
  const found: Array<{ pos: number; label: string }> = []
  block.forEach((child, offset) => {
    if (child.type.name !== 'field') return
    const instr = String(child.attrs['instr'] ?? '')
    if (fieldKind(instr) === 'SEQ') found.push({ pos: pos + 1 + offset, label: fieldArgument(instr) ?? '' })
  })
  return found
}

/** Os que o documento já usa e os do Word (Figura, Tabela, Equação) no idioma da interface. */
export function captionLabels(doc: ProseMirrorNode, defaults: readonly string[]): string[] {
  const labels = new Set(defaults)
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    for (const sequence of sequencesIn(node, pos)) if (sequence.label !== '') labels.add(sequence.label)
    return false
  })
  return [...labels]
}

export interface CaptionRequest {
  readonly label: string
  /** Pode ser vazio. */
  readonly text: string
  /** Acima é o costume das tabelas; abaixo, o das figuras. */
  readonly above: boolean
}

/**
 * No estilo `caption`, com um `SEQ` já contado, como no Word. As legendas de
 * depois só se renumeram com "Atualizar campos", também como no Word.
 */
export function insertCaption(editor: Editor, context: ReferenceContext, request: CaptionRequest): void {
  flushSelection(editor)
  const { state } = editor
  const ensured = ensureCaptionStyle(context.styles)
  const label = request.label.trim()
  if (label === '') return

  const $from = state.selection.$from
  const top = $from.depth === 0 ? $from.pos : $from.before(1)
  const block = state.doc.nodeAt(top)
  const at = request.above || block === null ? top : top + block.nodeSize

  const text = request.text.trim()
  const paragraph = state.schema.nodeFromJSON({
    type: 'paragraph',
    attrs: { styleId: ensured.id },
    content: [
      { type: 'text', text: `${label} ` },
      { type: 'field', attrs: { instr: ` SEQ ${label} \\* ARABIC `, result: '1' } },
      ...(text === '' ? [] : [{ type: 'text', text: ` ${text}` }]),
    ],
  })

  const tr = state.tr.insert(at, paragraph)
  if (ensured.sheet !== context.styles) context.setStyles(ensured.sheet)
  editor.view.dispatch(tr.scrollIntoView())

  const fieldPos = at + 1 + label.length + 1
  updateFieldsIn(editor, context, fieldPos, fieldPos + 1, new Set(['SEQ']))
}

/** A que a referência aponta: um título, um marcador, ou a legenda de um rótulo. */
export type CrossReferenceKind =
  | { readonly type: 'heading' }
  | { readonly type: 'bookmark' }
  | { readonly type: 'caption'; readonly label: string }
  | { readonly type: 'note'; readonly kind: NoteKind }

export interface CrossReferenceTarget {
  readonly key: string
  readonly text: string
}

export function crossReferenceTargets(
  doc: ProseMirrorNode,
  sheet: StyleSheet,
  kind: CrossReferenceKind,
  /** Para listar as notas pelo número da tela (`noteLabelsOf`). */
  labels?: readonly string[],
): CrossReferenceTarget[] {
  if (kind.type === 'heading') {
    return outlineOf(outlineBlocksOf(doc), sheet).map((entry) => ({
      key: String(entry.pos),
      text: `${' '.repeat(entry.level - 1)}${entry.text}`,
    }))
  }

  if (kind.type === 'bookmark') {
    return bookmarksOf(doc)
      .filter((bookmark) => !bookmark.name.startsWith('_'))
      .map((bookmark) => ({ key: bookmark.name, text: bookmark.name }))
  }

  if (kind.type === 'note') {
    // O número e o começo do texto, como o Word lista.
    return noteRefsOf(doc).flatMap(({ node, pos }, index) => {
      if (node.attrs['kind'] !== kind.kind) return []
      const text = node.textBetween(0, node.content.size, ' ').trim()
      const label = labels?.[index] ?? ''
      const short = text.length > 60 ? `${text.slice(0, 60)}…` : text
      return [{ key: String(pos), text: `${label} ${short}`.trim() }]
    })
  }

  const wanted = kind.label.toLowerCase()
  const captions: CrossReferenceTarget[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (sequencesIn(node, pos).some((sequence) => sequence.label.toLowerCase() === wanted)) {
      captions.push({ key: String(pos), text: textBetween(doc, pos + 1, pos + node.nodeSize - 1).trim() })
    }
    return false
  })
  return captions
}

export type CrossReferenceShow = 'text' | 'number' | 'page'

export interface CrossReferenceRequest {
  readonly kind: CrossReferenceKind
  readonly key: string
  readonly show: CrossReferenceShow
  /** `\h`: com Ctrl+clique, a referência leva ao destino. */
  readonly link: boolean
}

/**
 * O existente vale quando termina logo depois do trecho e começa até `slack`
 * posições antes: o marcador da legenda do Word começa antes do "Figura".
 */
function rangeBookmark(tr: Transaction, from: number, to: number, slack: number): string {
  const existing = bookmarksOf(tr.doc).find(
    (bookmark) =>
      bookmark.name.startsWith('_Ref') &&
      bookmark.end === to &&
      bookmark.pos < from &&
      bookmark.pos >= from - 1 - slack,
  )
  if (existing !== undefined) return existing.name

  const ids: string[] = []
  tr.doc.descendants((node) => {
    if (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd')
      ids.push(String(node.attrs['bid']))
    return true
  })
  const name = hiddenBookmarkName(
    '_Ref',
    bookmarksOf(tr.doc).map((bookmark) => bookmark.name),
  )
  const bid = nextBookmarkId(ids)
  const schema = tr.doc.type.schema
  tr.insert(to, schema.nodes['bookmarkEnd']!.create({ bid }))
  tr.insert(from, schema.nodes['bookmarkStart']!.create({ name, bid }))
  return name
}

/**
 * O título e a legenda ganham na hora o marcador oculto, como no Word. Da
 * legenda, "texto" é rótulo e número, e "número" só o número.
 */
export function insertCrossReference(
  editor: Editor,
  context: ReferenceContext,
  request: CrossReferenceRequest,
): boolean {
  flushSelection(editor)
  const { state } = editor
  const tr = state.tr
  let name: string | null

  if (request.kind.type === 'bookmark') {
    name = request.key
  } else if (request.kind.type === 'heading') {
    name = ensureBookmarks(tr, [Number(request.key)], '_Ref')[0] ?? null
  } else if (request.kind.type === 'note') {
    // O marcador que o `NOTEREF` cita, como o Word grava.
    const pos = Number(request.key)
    const reference = tr.doc.nodeAt(pos)
    if (reference === null || reference.type.name !== 'noteRef') return false
    name = rangeBookmark(tr, pos, pos + reference.nodeSize, 0)
  } else {
    const pos = Number(request.key)
    const block = tr.doc.nodeAt(pos)
    const wanted = request.kind.label.toLowerCase()
    const sequence =
      block === null ? undefined : sequencesIn(block, pos).find((item) => item.label.toLowerCase() === wanted)
    if (sequence === undefined) return false
    const first = textStartOf(tr.doc, pos)
    // Um marcador por legenda, para o texto, o número (`\# 0`) e a página, como o Word.
    name = rangeBookmark(tr, first, sequence.pos + 1, first - pos - 1)
  }
  if (name === null) return false

  const switches = `${request.show === 'number' ? ' \\# 0' : ''}${request.link ? ' \\h' : ''}`
  const instr =
    request.show === 'page'
      ? ` PAGEREF ${name}${switches} `
      : request.kind.type === 'note'
        ? ` NOTEREF ${name}${request.link ? ' \\h' : ''} `
        : ` REF ${name}${switches} `
  const at = tr.mapping.map(state.selection.from)
  tr.replaceWith(
    at,
    tr.mapping.map(state.selection.to),
    state.schema.nodes['field']!.create({ instr, result: '' }),
  )
  editor.view.dispatch(tr.scrollIntoView())

  // Pelo mesmo caminho do F9.
  updateFieldsIn(editor, context, at, at + 1)
  return true
}
