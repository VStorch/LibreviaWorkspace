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
import { outlineOf } from '@services/document/outline.js'
import {
  ensureCaptionStyle,
  ensureTocHeadingStyle,
  ensureTocStyle,
} from '@services/document/reference-styles.js'
import type { StyleSheet } from '@services/document/styles.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { bookmarksOf } from './extensions/bookmark.js'
import { DEFAULT_TOC_INSTRUCTION } from './extensions/table-of-contents.js'
import { outlineBlocksOf } from './outline-blocks.js'
import type { PageLayout, PageStart } from './usePagination.js'

/**
 * O que as referências precisam saber além do documento: em que folha cada coisa
 * caiu (a paginação do M6), como a folha numera (M7) e os estilos.
 *
 * Lido na hora do comando, e não guardado: a paginação muda a cada linha
 * digitada, e um número de página calculado com a de antes é o erro que o
 * sumário existe para não ter.
 */
export interface ReferenceContext {
  readonly layout: PageLayout
  readonly page: PageSetup
  readonly styles: StyleSheet
  readonly setStyles: (styles: StyleSheet) => void
  readonly t: (key: MessageKey) => string
  /** Marcadores que existem no arquivo fora dos nós — ver `DocumentModel.outsideBookmarks`. */
  readonly outsideBookmarks?: readonly string[]
}

/** Faz o ProseMirror ler agora a seleção que o navegador já mudou. */
export function flushSelection(editor: Editor): void {
  ;(editor.view as { domObserver?: { flush?: () => void } }).domObserver?.flush?.()
}

// --- página de uma posição -------------------------------------------------

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

// --- campos ------------------------------------------------------------------

/** O texto entre duas posições, com o resultado dos campos no lugar deles. */
function textBetween(doc: ProseMirrorNode, from: number, to: number): string {
  return doc.textBetween(from, to, ' ', (leaf) =>
    leaf.type.name === 'field' ? String(leaf.attrs['result'] ?? '') : '',
  )
}

/** Os tipos que dependem de onde o texto cai na folha. */
const PAGE_KINDS = new Set(['PAGE', 'PAGEREF', 'NUMPAGES'])

/** Os tipos que o editor sabe recalcular. O resto fica como o Word o deixou. */
const UPDATABLE = new Set(['PAGE', 'PAGEREF', 'NUMPAGES', 'REF', 'SEQ'])

export interface FieldUpdate {
  /** Quantos campos mudaram de resultado. */
  readonly changed: number
  /** Algum campo de página foi recalculado — e a paginação pode mudar com ele. */
  readonly pageDependent: boolean
}

/**
 * Recalcula os campos entre `from` e `to` — o F9 do Word.
 *
 * A ordem é a do Word: primeiro as sequências (`SEQ`), porque a referência a uma
 * legenda (`REF`) cita o número dela; depois as referências; e as páginas por
 * último, com a paginação que a tela tem agora. O `SEQ` conta o documento
 * inteiro — o número de uma legenda depende das de antes, estejam ou não na
 * seleção —, mas só os de dentro dela mudam.
 *
 * O campo que aponta para marcador que não existe mais ganha o texto do Word,
 * "Erro! Indicador não definido.", e não some: é assim que a pessoa descobre que
 * apagou o que ele citava.
 */
export function updateFieldsIn(
  editor: Editor,
  context: ReferenceContext,
  from: number,
  to: number,
  kinds: ReadonlySet<string> = UPDATABLE,
): FieldUpdate {
  // O cursor que a pessoa acabou de mover pode estar só no DOM; a transação
  // abaixo leva a seleção do estado, e sem isto o devolveria ao lugar de antes —
  // o segundo passe chega sozinho, a qualquer momento, e puxava o cursor de volta.
  flushSelection(editor)
  const { state } = editor
  const doc = state.doc

  const fields: Array<{ pos: number; node: ProseMirrorNode; kind: string }> = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'field') fields.push({ pos, node, kind: fieldKind(String(node.attrs['instr'])) })
    return true
  })

  // As sequências contam o documento inteiro, na ordem.
  const sequences = fields.filter((field) => field.kind === 'SEQ')
  const numbers = sequenceNumbers(sequences.map((field) => String(field.node.attrs['instr'])))
  const sequenceResult = new Map(sequences.map((field, index) => [field.pos, numbers[index]!]))

  // A referência lê o texto do marcador **depois** de as sequências mudarem: a
  // legenda "Figura 1" que virou "Figura 2" tem de ser citada como "Figura 2".
  const updatedSequenceDoc = (() => {
    const tr = state.tr
    for (const [pos, result] of sequenceResult) tr.setNodeAttribute(pos, 'result', result)
    return tr.doc
  })()

  const bookmarks = new Map(bookmarksOf(updatedSequenceDoc).map((bookmark) => [bookmark.name, bookmark]))
  const missing = context.t('references.field.missingBookmark')
  // O marcador que o arquivo tem fora dos nós não está perdido: a referência a
  // ele fica com o resultado que o Word calculou.
  const outside = new Set(context.outsideBookmarks ?? [])
  const sheets = context.layout.pages

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
        // `\# 0`: só o número do texto citado — "Figura 2" vira "2". É como o
        // Word faz a referência "só o número" a uma legenda.
        result = fieldSwitch(instr, '#') === null ? text : (/(\d+)(?!.*\d)/.exec(text)?.[1] ?? text)
        break
      }
      case 'PAGEREF': {
        const name = fieldArgument(instr) ?? ''
        const target = bookmarks.get(name)
        if (target === undefined) {
          result = outside.has(name) ? null : missing
          break
        }
        result = pageLabel(context.page, sheetAt(doc, context.layout.pageStarts, target.pos))
        pageDependent = true
        break
      }
      case 'PAGE':
        result = pageLabel(context.page, sheetAt(doc, context.layout.pageStarts, field.pos))
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
 * Os campos da seleção, ou do documento inteiro com o cursor parado — e, se
 * algum deles depende da página, um segundo passe quando a paginação assentar.
 *
 * O segundo passe é o que faz o número da página sair certo: atualizar muda o
 * texto dos campos, o texto pode empurrar uma linha para a folha seguinte, e o
 * número calculado antes do empurrão fica velho. O Word também faz os dois.
 */
export function updateFields(editor: Editor, context: ReferenceContext): FieldUpdate {
  const { from, to, empty } = editor.state.selection
  const range = empty ? { from: 0, to: editor.state.doc.content.size } : { from, to }
  const update = updateFieldsIn(editor, context, range.from, range.to)
  // Só quando algo mudou: armado à toa, o passe dispararia na próxima digitação
  // e reescreveria campos que ninguém mandou atualizar.
  if (update.pageDependent && update.changed > 0 && empty) arm(editor, 'all')
  return update
}

/**
 * Quem pediu o segundo passe: o F9 do documento inteiro corrige todos os campos
 * de página; o sumário, só os dele — o resto do documento a pessoa não mandou
 * atualizar.
 */
const pendingPagePass = new WeakMap<Editor, { scope: 'all' | 'toc'; doc: ProseMirrorNode }>()

/** Arma o segundo passe para o documento de agora. */
function arm(editor: Editor, scope: 'all' | 'toc'): void {
  pendingPagePass.set(editor, { scope, doc: editor.state.doc })
}

/**
 * O segundo passe, chamado quando a paginação muda. Só os campos de página, e uma
 * vez: o passe não pede outro.
 */
export function settlePageFields(editor: Editor, context: ReferenceContext): void {
  const pending = pendingPagePass.get(editor)
  if (pending === undefined) return
  pendingPagePass.delete(editor)
  // A pessoa voltou a escrever antes de a paginação assentar: o passe é de
  // outro documento, e não vale mais.
  if (pending.doc !== editor.state.doc) return
  const { scope } = pending

  if (scope === 'all') {
    updateFieldsIn(editor, context, 0, editor.state.doc.content.size, PAGE_KINDS)
    return
  }

  // De trás para a frente não é preciso: mudar o resultado não desloca posição.
  for (const { pos, node } of tablesOfContents(editor.state.doc)) {
    updateFieldsIn(editor, context, pos, pos + node.nodeSize, PAGE_KINDS)
  }
}

// --- marcadores ocultos dos títulos -----------------------------------------------

/**
 * Garante o marcador oculto em volta do texto de cada bloco dado, na mesma
 * transação, e devolve o nome de cada um na ordem dada.
 *
 * De trás para a frente, para que a inserção de um não desloque a posição dos
 * que ainda faltam.
 */
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

// --- sumário ---------------------------------------------------------------

/** As posições dos sumários do documento. */
function tablesOfContents(doc: ProseMirrorNode): Array<{ pos: number; node: ProseMirrorNode }> {
  const found: Array<{ pos: number; node: ProseMirrorNode }> = []
  doc.forEach((node, pos) => {
    if (node.type.name === 'tableOfContents') found.push({ pos, node })
  })
  return found
}

/**
 * Monta as entradas de um sumário na transação: marca os títulos com `_Toc…` e
 * devolve os parágrafos, na forma do Word — o texto do título e uma tabulação,
 * e o `PAGEREF` do número, os dois dentro do link para o marcador.
 */
function buildEntries(
  tr: Transaction,
  context: ReferenceContext,
  instr: string,
  sheet: StyleSheet,
): { entries: unknown[]; sheet: StyleSheet } {
  const { from, to } = tocLevels(instr)
  const links = tocLinks(instr)
  const omitPages = tocOmitsPages(instr)

  // Os títulos de fora dos sumários: a entrada de um sumário nunca é título, mas
  // um sumário antigo com parágrafo em estilo de título se listaria a si mesmo.
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
    // A folha é lida na paginação que a tela tem agora, e por isso no documento
    // de antes dos marcadores novos, que é o que ela mediu. O segundo passe (ver
    // `settlePageFields`) a corrige depois de o sumário ocupar o lugar dele.
    const page = pageLabel(context.page, sheetAt(tr.before, context.layout.pageStarts, heading.pos))
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

/**
 * Insere um sumário no lugar do bloco do cursor — antes dele, ou no lugar dele se
 * estiver vazio —, com o título "Sumário" e as entradas dos títulos 1 a 3.
 */
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

  // O bloco de primeiro nível do cursor, na transação que já tem os marcadores.
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
 * Refaz as entradas do sumário do cursor — ou do primeiro, com o cursor fora
 * de todos —, a partir dos títulos de agora.
 *
 * O título do sumário (os parágrafos antes de `head`), a instrução e o controle
 * de conteúdo ficam; as entradas são trocadas inteiras, como no "Atualizar
 * sumário inteiro" do Word. A correção feita à mão numa entrada vai embora, e lá
 * também.
 *
 * Devolve falso quando o documento não tem sumário.
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

  // A posição do sumário depois dos marcadores novos, que podem ter entrado antes
  // dele (um título antes do sumário é raro, mas existe).
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

// --- legendas ----------------------------------------------------------------

/** Os campos `SEQ` de um bloco de texto, com a posição de cada um. */
function sequencesIn(block: ProseMirrorNode, pos: number): Array<{ pos: number; label: string }> {
  const found: Array<{ pos: number; label: string }> = []
  block.forEach((child, offset) => {
    if (child.type.name !== 'field') return
    const instr = String(child.attrs['instr'] ?? '')
    if (fieldKind(instr) === 'SEQ') found.push({ pos: pos + 1 + offset, label: fieldArgument(instr) ?? '' })
  })
  return found
}

/**
 * Os rótulos de legenda: os que o documento já usa em algum `SEQ`, e os do Word
 * (Figura, Tabela, Equação) no idioma da interface.
 */
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
  /** O que vem depois do número: "— Arquitetura do sistema". Pode ser vazio. */
  readonly text: string
  /** Acima do bloco do cursor (o costume das tabelas) ou abaixo (o das figuras). */
  readonly above: boolean
}

/**
 * Insere uma legenda — "Figura 1 — texto" — num parágrafo novo, acima ou abaixo
 * do bloco do cursor, no estilo `caption`.
 *
 * O número é um `SEQ` do rótulo, como no Word, e sai já contado: as legendas
 * de antes dela dão o número. As de depois não mudam sozinhas; "Atualizar
 * campos" as renumera, também como no Word.
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

  // O número certo pela contagem do documento — só o campo novo muda.
  const fieldPos = at + 1 + label.length + 1
  updateFieldsIn(editor, context, fieldPos, fieldPos + 1, new Set(['SEQ']))
}

// --- referências cruzadas -----------------------------------------------------

/** A que a referência aponta: um título, um marcador, ou a legenda de um rótulo. */
export type CrossReferenceKind =
  | { readonly type: 'heading' }
  | { readonly type: 'bookmark' }
  | { readonly type: 'caption'; readonly label: string }

export interface CrossReferenceTarget {
  /** Posição do bloco (título, legenda) ou nome do marcador. */
  readonly key: string
  readonly text: string
}

/** Os destinos possíveis do tipo escolhido, na ordem do documento. */
export function crossReferenceTargets(
  doc: ProseMirrorNode,
  sheet: StyleSheet,
  kind: CrossReferenceKind,
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

/** O que a referência mostra: o texto do destino, o número da legenda, ou a página. */
export type CrossReferenceShow = 'text' | 'number' | 'page'

export interface CrossReferenceRequest {
  readonly kind: CrossReferenceKind
  readonly key: string
  readonly show: CrossReferenceShow
  /** `\h`: com Ctrl+clique, a referência leva ao destino. */
  readonly link: boolean
}

/**
 * O marcador oculto que cobre o trecho `[from, to)` — o que já existe, ou um
 * `_Ref` novo — e devolve o nome.
 *
 * O que já existe vale quando termina logo depois do trecho e começa no trecho ou
 * até `slack` posições antes dele: o marcador da legenda que o Word grava começa
 * antes do "Figura", e outros marcadores podem estar colados no começo.
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
 * Insere uma referência cruzada no cursor: um `REF` (o texto ou o número) ou um
 * `PAGEREF` (a página) para o marcador do destino — que o título e a legenda
 * ganham na hora, oculto, como no Word.
 *
 * Da legenda, "texto" é o rótulo e o número ("Figura 1"), e "número" só o número:
 * são dois marcadores, um em volta de cada trecho.
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
  } else {
    const pos = Number(request.key)
    const block = tr.doc.nodeAt(pos)
    const wanted = request.kind.label.toLowerCase()
    const sequence =
      block === null ? undefined : sequencesIn(block, pos).find((item) => item.label.toLowerCase() === wanted)
    if (sequence === undefined) return false
    // O começo do texto da legenda: depois dos marcadores que abrem o parágrafo.
    let first = pos + 1
    for (let index = 0; block !== null && index < block.childCount; index++) {
      const child = block.child(index)
      if (child.type.name !== 'bookmarkStart' && child.type.name !== 'bookmarkEnd') break
      first += child.nodeSize
    }
    // Um marcador só por legenda, em volta do rótulo e do número, para o texto,
    // o número (`\# 0`) e a página — como o Word.
    name = rangeBookmark(tr, first, sequence.pos + 1, first - pos - 1)
  }
  if (name === null) return false

  const switches = `${request.show === 'number' ? ' \\# 0' : ''}${request.link ? ' \\h' : ''}`
  const instr = request.show === 'page' ? ` PAGEREF ${name}${switches} ` : ` REF ${name}${switches} `
  const at = tr.mapping.map(state.selection.from)
  tr.replaceWith(
    at,
    tr.mapping.map(state.selection.to),
    state.schema.nodes['field']!.create({ instr, result: '' }),
  )
  editor.view.dispatch(tr.scrollIntoView())

  // O resultado sai do cálculo, e não de uma cópia: é o mesmo caminho do F9.
  updateFieldsIn(editor, context, at, at + 1)
  return true
}
