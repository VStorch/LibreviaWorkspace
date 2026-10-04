import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import {
  noteLineTop,
  noteSpan,
  paginateSections,
  type MeasuredBlock,
  type MeasuredNote,
  type NoteSlice,
  type SectionFlow,
  type SheetPlan,
} from '@services/document/paginate.js'
import { NoteKind } from '@services/document/notes.js'
import { effectiveAttrs } from '@services/document/style-cascade.js'
import type { StyleSheet } from '@services/document/styles.js'
import {
  contentHeightMm,
  contentInsetsMm,
  pageDimensionsMm,
  type PageSetup,
  type SectionSetup,
} from '@services/document/model.js'
import { mmToPx } from '@services/units.js'
import { NO_BANDS, type BandHeights } from '@services/document/band.js'
import {
  blockSections,
  parityOf,
  sectionBreakIn,
  startsNewSheet,
  columnGeometry,
  type SectionBlock,
} from '@services/document/sections.js'
import { applyPageGaps, type RepeatedHeader } from './extensions/pagination.js'
import { LINE_GAP_CLASS, measureLines } from './line-boxes.js'
import { noteBodyOf, type NoteBody } from './extensions/note-view.js'

/** Espaço entre uma folha e a seguinte, como numa pilha de papel. */
export const SHEET_GUTTER_PX = 28

/** Uma linha de 12 pt com o traço no meio: o parágrafo do separador do Word. */
export const NOTE_SEPARATOR_PX = 16

export interface NoteAreaItem {
  /** O corpo na tela (`note-view.ts`). */
  readonly key: string
  /** A ordem da referência entre todas as do documento (`noteRefsOf`) — o papel acha o nó por ela. */
  readonly index: number
  readonly fromLine: number
  readonly toLine: number
  readonly clipTopPx: number
  readonly heightPx: number
}

/**
 * A área de notas de uma folha: as de rodapé no pé da coluna de texto, as de
 * fim logo depois do último bloco (e nas folhas que vierem depois dele).
 */
export interface NoteArea {
  readonly sheet: number
  readonly kind: 'footnote' | 'endnote'
  /** Em pixels da folha, já com o separador. */
  readonly topPx: number
  readonly leftPx: number
  readonly widthPx: number
  readonly separator: 'normal' | 'continuation' | null
  readonly items: readonly NoteAreaItem[]
}

/** Constante, para `sameGaps` comparar por valor. */
const EMPTY_GAPS = new Map<number, number>()

export interface PageLayout {
  /** Quantas folhas desenhar — as em branco das seções par e ímpar incluídas. */
  readonly pages: number
  readonly stackHeightPx: number
  readonly sheetTops: readonly number[]
  readonly sheetHeights: readonly number[]
  /** Recortar a lista de blocos nestes pontos dá ao papel as mesmas folhas da tela. */
  readonly pageStarts: readonly PageStart[]
  /** A posição de um objeto ancorado vem do parágrafo, que só tem posição depois de paginar. */
  readonly anchors: readonly BlockAnchor[]
  readonly sheets: readonly SheetPlan[]
  readonly sheetWidths: readonly number[]
  /** A da folha mais larga; as outras vão centradas. */
  readonly stackWidthPx: number
  /** `pageStarts` conta só as folhas com conteúdo; as em branco ficam entre elas. */
  readonly contentSheets: readonly number[]
  /** O papel repete o mesmo desvio (`print-source.ts`). */
  readonly columnMoves: readonly ColumnMove[]
  readonly columnLines: readonly ColumnLine[]
  readonly noteAreas: readonly NoteArea[]
}

/** Um bloco de seção com colunas: o lado da coluna e o quanto subiu ou desceu. */
export interface ColumnMove {
  readonly blockIndex: number
  readonly dx: number
  /** Quanto a coluna é mais estreita que a coluna de texto da folha. */
  readonly narrowerPx: number
  readonly lift: number
  /** O papel escreve margem natural mais desvio. */
  readonly natural: number
  /** A margem de baixo do bloco anterior — ver `collapsed`. */
  readonly collapse: number
}

/**
 * As margens verticais colapsam: positiva com positiva vale a maior, e uma
 * negativa se **soma** à positiva. Por isso a distância menor que a margem de
 * baixo do anterior só se obtém descontando-a.
 */
export function collapsed(distance: number, previous: number): number {
  return distance >= previous ? distance : distance - previous
}

/** Uma linha entre colunas, em pixels da folha desenhada. */
export interface ColumnLine {
  readonly sheet: number
  readonly leftPx: number
  readonly topPx: number
  readonly heightPx: number
}

/** A folha desenhada em que cai a folha de conteúdo `index`. */
export function drawnSheet(layout: PageLayout, index: number): number {
  return layout.contentSheets[index] ?? index
}

interface SectionMetrics {
  readonly columns: number
  readonly columnStepPx: number
  readonly columnWidthPx: number
  readonly separator: boolean
  readonly leftPx: number
  readonly rightPx: number
  readonly widthPx: number
  readonly heightPx: number
  readonly contentPx: number
  readonly topPx: number
  readonly bottomPx: number
}

export interface PageStart {
  readonly blockIndex: number
  /** Índice da linha ou item que abre a folha, quando o corte é interno. */
  readonly childIndex?: number
  /** Posição dentro do conteúdo do parágrafo, a mesma que `Node.cut` recebe. */
  readonly offset?: number
  /** A folha abre com as linhas de cabeçalho da tabela repetidas. */
  readonly repeatHeader?: boolean
}

/** O corte cai dentro do bloco — o bloco começa na folha anterior. */
export function isInternalStart(start: PageStart): boolean {
  return start.childIndex !== undefined || start.offset !== undefined
}

interface CutTarget {
  readonly at: number
  readonly start: PageStart
  readonly nodes: readonly { position: number; natural: number; collapse?: number }[]
  /** Resolvida só se o corte for escolhido. */
  readonly line?: { readonly resolve: () => number | null; readonly block: number }
  /** Corte entre linhas de tabela com cabeçalho: o que se repete no alto da folha. */
  readonly header?: () => RepeatedHeader
}

export interface BlockAnchor {
  readonly pageIndex: number
  /** Topo do bloco dentro da folha, em pixels, já incluída a margem superior. */
  readonly topPx: number
}

/** O que o editor passa à paginação além do documento e das seções. */
export interface PaginationOptions {
  readonly bands?: readonly BandHeights[]
  /**
   * O modo de leitura desliga só o empurrão no DOM: a conta continua, e imprimir
   * de dentro dele sai com as mesmas folhas, porque as coordenadas de fluxo não
   * dependem dos vãos aplicados.
   */
  readonly paginated?: boolean
  /** O "manter com o próximo" pode vir do estilo, e o bloco só carrega o que o parágrafo declara. */
  readonly styles?: StyleSheet | null
}

/**
 * A medição é convertida em **coordenadas de fluxo** antes de decidir: o
 * `offsetTop` já inclui os vãos aplicados, e subtraí-los devolve a altura da
 * tira contínua. Aplicar o resultado não muda a entrada da próxima medida, e o
 * laço assenta.
 */
export function usePagination(
  editor: Editor | null,
  /** Com as faixas herdadas resolvidas (`effectiveSections`); a última é a do corpo. */
  sections: readonly PageSetup[],
  /** É pelo id que o bloco acha a sua seção. */
  declared: readonly SectionSetup[],
  revision: number,
  { bands = [], paginated = true, styles = null }: PaginationOptions = {},
): PageLayout {
  const [layout, setLayout] = useState<PageLayout>({
    pages: 1,
    stackHeightPx: 0,
    sheetTops: [0],
    sheetHeights: [],
    pageStarts: [],
    anchors: [],
    sheets: [{ section: 0, blank: false, number: 1, first: true }],
    sheetWidths: [],
    stackWidthPx: 0,
    contentSheets: [0],
    columnMoves: [],
    columnLines: [],
    noteAreas: [],
  })

  /**
   * Numa `ref`, porque o efeito é refeito a cada tecla e a leitura seguinte
   * precisa descontar o que já foi empurrado. A chave é a **posição** do nó: um
   * corte interno empurra linha de tabela ou item de lista, sem índice de bloco.
   */
  const applied = useRef(new Map<number, number>())

  /** O vão **mais** a margem natural, que é o que o CSS lê; o vão é o que a conta desconta. */
  const lastWritten = useRef(new Map<number, number>())

  /** Os vãos entre linhas de parágrafo cortado, por posição do espaçador. */
  const lastLines = useRef(new Map<number, number>())

  /** O deslocamento lateral das colunas já aplicado, por posição do bloco. */
  const lastColumns = useRef(new Map<number, number>())

  /** Os cabeçalhos de tabela repetidos, comparados pelo que desenham. */
  const lastHeaders = useRef('[]')

  const bandsKey = bands.map((band) => `${band.headerMm}:${band.footerMm}`).join('|')

  useEffect(() => {
    if (editor === null) return undefined

    const element = editor.view.dom as HTMLElement // alvo do observador de tamanho
    // A margem é um piso: o cabeçalho mais alto a empurra para baixo.
    const metrics: SectionMetrics[] = sections.map((setup, index) => {
      const heights = bands[index] ?? NO_BANDS
      const insets = contentInsetsMm(setup, heights)
      const { width, height } = pageDimensionsMm(setup)
      const columns = columnGeometry(setup)
      return {
        columns: columns.count,
        columnStepPx: mmToPx(columns.stepMm),
        columnWidthPx: mmToPx(columns.widthMm),
        separator: columns.separator,
        leftPx: mmToPx(setup.margins.left),
        rightPx: mmToPx(setup.margins.right),
        widthPx: mmToPx(width),
        heightPx: mmToPx(height),
        contentPx: mmToPx(contentHeightMm(setup, heights)),
        topPx: mmToPx(insets.top),
        bottomPx: mmToPx(insets.bottom),
      }
    })
    const metricsOf = (section: number): SectionMetrics => metrics[section] ?? metrics.at(-1)!
    const flows: SectionFlow[] = sections.map((setup, index) => ({
      height: metricsOf(index).contentPx,
      newSheet: startsNewSheet(setup, sections[index - 1]),
      parity: parityOf(setup),
      restart: setup.pageNumberStart ?? null,
      columns: metricsOf(index).columns,
    }))

    const measure = (): void => {
      // Os espaçadores entre linhas saem durante a medida: eles mudam onde as
      // linhas quebram, e a folha ficaria com linhas vazias no pé. O estilo volta
      // no mesmo quadro, antes de o navegador desenhar.
      const lineGapsInDom = Array.from(element.querySelectorAll<HTMLElement>(`.${LINE_GAP_CLASS}`))
      for (const gap of lineGapsInDom) gap.style.display = 'none'
      try {
        measureHidden()
      } finally {
        for (const gap of lineGapsInDom) gap.style.display = 'inline-block'
      }
    }

    const measureHidden = (): void => {
      // Pelo **documento**, e não pelos filhos do DOM, que têm outros índices:
      // `nodeDOM` liga um ao outro.
      let accumulated = 0
      const blocks: MeasuredBlock[] = []
      const targets: CutTarget[] = []
      const origin = offsetTopOf(element)

      // As notas: as de rodapé vão com o bloco da referência; as de fim, depois
      // do último bloco. `refIndex` é a ordem de `noteRefsOf`.
      const measuredNotes = new Map<string, MeasuredNote & { index: number }>()
      const endnotes: (MeasuredNote & { index: number })[] = []
      let refIndex = 0

      const marks: (string | null)[] = []
      editor.state.doc.forEach((block) => marks.push(sectionBreakIn(block as unknown as SectionBlock)))
      const sectionOfBlock = blockSections(marks, declared)

      editor.state.doc.forEach((block, offset, blockIndex) => {
        const section = sectionOfBlock[blockIndex] ?? 0
        const dom = editor.view.nodeDOM(offset)
        const node = dom instanceof HTMLElement ? dom : null
        if (node === null) {
          block.descendants((child) => {
            if (child.type.name !== 'noteRef') return true
            refIndex += 1
            return false
          })
          blocks.push({
            top: 0,
            height: 0,
            breakpoints: [],
            isPageBreak: false,
            breakAfter: false,
            keepWithNext: false,
            section,
          })
          return
        }

        accumulated += shiftOf(node)
        const top = offsetTopOf(node) - origin - accumulated
        const before = blocks.at(-1)
        const previousDom = node.previousElementSibling
        targets.push({
          at: top,
          start: { blockIndex },
          nodes: [
            {
              position: offset,
              natural: Math.max(top - (before === undefined ? 0 : before.top + before.height), 0),
              // É com ela que uma margem de cima negativa se soma (`collapsed`).
              collapse:
                previousDom instanceof HTMLElement
                  ? parseFloat(getComputedStyle(previousDom).marginBottom) || 0
                  : 0,
            },
          ],
        })

        // Só as linhas da tabela externa: as aninhadas pertencem às células.
        const table =
          node instanceof HTMLTableElement ? node : node.querySelector<HTMLTableElement>(':scope > table')
        const children =
          table !== null
            ? Array.from(table.rows)
            : node.tagName === 'UL' || node.tagName === 'OL'
              ? Array.from(node.children).filter(
                  (child): child is HTMLElement => child instanceof HTMLElement && child.tagName === 'LI',
                )
              : // O sumário corta entre entradas, como a lista entre itens.
                node.hasAttribute('data-toc')
                ? Array.from(node.children).filter(
                    (child): child is HTMLElement => child instanceof HTMLElement,
                  )
                : []
        let internal = 0
        const breakpoints: number[] = []
        // O pé da linha da referência de nota que está dentro dela.
        const childTops: number[] = []

        // Parágrafo e título cortam entre linhas. As linhas medem a partir da
        // borda do bloco, e o topo de fluxo dele já está em `top`.
        const lines = block.isTextblock && children.length === 0 ? measureLines(editor.view, node) : null
        if (lines !== null) {
          lines.starts.forEach((start, index) => {
            const at = top + start
            breakpoints.push(at)
            targets.push({
              at,
              start: { blockIndex },
              nodes: [],
              line: { resolve: () => lines.positionOf(index), block: offset },
            })
          })
          internal = lines.shift
        }

        // A captura ancorada: a linha vazia depois do quadro (`::after` de 1lh)
        // passa para a folha seguinte quando não cabe, como no LibreOffice.
        const freeBreakpoints: number[] = []
        let hangingBottom = 0
        if (
          lines !== null &&
          node.querySelector(':scope > .node-image[data-anchored], :scope > img[data-anchored]') !== null
        ) {
          // Com texto, o corte entre o quadro e a linha é livre da regra de viúvas.
          const first = lines.starts[0]
          if (
            first !== undefined &&
            block.firstChild?.type.name === 'image' &&
            block.firstChild.attrs['anchored'] === true
          ) {
            freeBreakpoints.push(top + first)
          }
          // Sem texto, a linha vazia pode sobrar no pé da folha.
          const after = parseFloat(getComputedStyle(node, '::after').height)
          if (Number.isFinite(after) && after > 0) hangingBottom = after
        }

        // Linhas de cabeçalho (`w:tblHeader`) no começo da tabela: cortar dentro
        // delas, ou logo depois, deixaria o cabeçalho sozinho no pé.
        const rows = table !== null ? Array.from(table.rows) : []
        let headerRows = 0
        while (
          headerRows < rows.length - 1 &&
          rows[headerRows]!.cells.length > 0 &&
          Array.from(rows[headerRows]!.cells).every((cell) => cell.tagName === 'TH')
        ) {
          headerRows += 1
        }
        const lastHeader = rows[headerRows - 1]
        const repeatHeight =
          lastHeader === undefined
            ? 0
            : offsetTopOf(lastHeader) + lastHeader.offsetHeight - offsetTopOf(rows[0]!)

        children.forEach((child, childIndex) => {
          const cells = child instanceof HTMLTableRowElement ? Array.from(child.cells) : []
          const shift = cells.length > 0 ? shiftOf(cells[0]!) : shiftOf(child)
          // Padding aumenta a linha para baixo; margem já deslocou seu topo.
          const at = offsetTopOf(child) - origin - accumulated - internal - (cells.length === 0 ? shift : 0)
          childTops.push(at)
          if (childIndex > headerRows) {
            breakpoints.push(at)
            targets.push({
              at,
              start: { blockIndex, childIndex },
              nodes: (cells.length > 0 ? cells : [child]).map((target) => ({
                position: editor.view.posAtDOM(target, 0) - 1,
                natural:
                  parseFloat(
                    cells.length > 0
                      ? getComputedStyle(target).paddingTop
                      : getComputedStyle(target).marginTop,
                  ) - shiftOf(target),
              })),
              ...(table !== null && repeatHeight > 0
                ? {
                    header: () =>
                      repeatedHeader(editor, table, rows.slice(0, headerRows), cells[0]!, repeatHeight),
                  }
                : {}),
            })
          }
          internal += shift
        })
        const effective = effectiveAttrs(block, styles)
        const height = node.offsetHeight - internal

        const blockNotes: MeasuredNote[] = []
        block.descendants((child, pos) => {
          if (child.type.name !== 'noteRef') return true
          const index = refIndex++
          const reference = editor.view.nodeDOM(offset + 1 + pos)
          const body = noteBodyOf(reference)
          if (body === undefined || !(reference instanceof HTMLElement)) return false
          const at = referenceBottom(node, reference, top, height, {
            lineStarts: lines?.starts ?? null,
            children,
            childTops,
          })
          const measured = { id: body.key, at, index, ...measureNote(body) }
          measuredNotes.set(body.key, measured)
          if (child.attrs['kind'] === NoteKind.Endnote) endnotes.push(measured)
          else blockNotes.push(measured)
          return false
        })

        blocks.push({
          top,
          height,
          breakpoints,
          isPageBreak: node.hasAttribute('data-page-break'),
          breakAfter: node.hasAttribute('data-break-after'),
          columnBreakAfter: node.hasAttribute('data-column-break'),
          keepWithNext: effective['keepNext'] === true || /^H[1-6]$/.test(node.tagName),
          keepLines: effective['keepLines'] === true,
          widowControl: lines !== null && effective['widowControl'] !== false,
          ...(repeatHeight > 0 ? { repeatHeight } : {}),
          ...(freeBreakpoints.length > 0 ? { freeBreakpoints } : {}),
          ...(hangingBottom > 0 ? { hangingBottom } : {}),
          ...(blockNotes.length > 0 ? { notes: blockNotes } : {}),
          section,
        })
        accumulated += internal
      })

      // As notas de fim entram no fluxo depois do último bloco, cortando entre linhas.
      const textBottom = blocks.reduce((bottom, block) => Math.max(bottom, block.top + block.height), 0)
      const lastSection = blocks.at(-1)?.section ?? 0
      const endnoteBlocks: MeasuredBlock[] = []
      let endnoteTop = textBottom + NOTE_SEPARATOR_PX
      for (const note of endnotes) {
        endnoteBlocks.push({
          top: endnoteTop,
          height: note.height,
          breakpoints: note.lines.slice(1).map((line) => endnoteTop + line),
          isPageBreak: false,
          breakAfter: false,
          keepWithNext: false,
          section: lastSection,
        })
        endnoteTop += note.height
      }

      const plan = paginateSections([...blocks, ...endnoteBlocks], flows, { separator: NOTE_SEPARATOR_PX })
      const breaks = plan.breaks
      const contentSheets = plan.sheets.flatMap((sheet, index) => (sheet.blank ? [] : [index]))
      const sheetOf = (content: number): SheetPlan => plan.sheets[contentSheets[content] ?? 0]!
      // A altura desenhada de uma folha é a da tira mais os desvios das colunas.
      const liftsBetween = (from: number, to: number): number => {
        let sum = 0
        for (const [index, placement] of plan.placements) {
          const top = blocks[index]?.top
          if (top !== undefined && top >= from && top < to) sum += placement.lift
        }
        return sum
      }

      // Vão = o que sobrou da folha + as duas margens + o espaço entre papéis.
      // A margem escrita **substitui** a margem natural do bloco (a última
      // declaração ganha), e por isso o valor escrito é o vão mais a margem
      // natural; a conta de fluxo desconta só o vão.
      const gaps = new Map<number, number>()
      const written = new Map<number, number>()
      const lineGaps = new Map<number, number>()
      const headers: RepeatedHeader[] = []
      let previous = 0
      const pageStarts: PageStart[] = []
      const sheetHeights: number[] = []
      const noteAreas: NoteArea[] = []
      const sheetStarts: number[] = []

      // No pé da coluna de texto, ou logo depois do texto se a folha esticou.
      const footnoteArea = (content: number, metrics: SectionMetrics, used: number): void => {
        const slices = plan.notes[content] ?? []
        const items = slices.flatMap((slice) => itemOf(slice))
        if (items.length === 0) return
        const height = plan.noteHeights[content] ?? 0
        noteAreas.push({
          sheet: contentSheets[content] ?? content,
          kind: 'footnote',
          topPx: metrics.topPx + Math.max(metrics.contentPx - height, used),
          leftPx: metrics.leftPx,
          widthPx: metrics.widthPx - metrics.leftPx - metrics.rightPx,
          separator: items[0]!.fromLine > 0 ? 'continuation' : 'normal',
          items,
        })
      }
      const itemOf = (slice: NoteSlice): NoteAreaItem[] => {
        const note = measuredNotes.get(slice.id)
        if (note === undefined) return []
        return [
          {
            key: slice.id,
            index: note.index,
            fromLine: slice.fromLine,
            toLine: slice.toLine,
            clipTopPx: noteLineTop(note, slice.fromLine),
            heightPx: noteSpan(note, slice.fromLine, slice.toLine),
          },
        ]
      }

      // A lista e os cortes crescem juntos: cada folha custa uma consulta.
      const internalAt = new Map<number, CutTarget>()
      for (const target of targets) {
        if (
          (target.start.childIndex !== undefined || target.line !== undefined) &&
          !internalAt.has(target.at)
        ) {
          internalAt.set(target.at, target)
        }
      }
      let cursor = 0
      const blockTargetFrom = (at: number): CutTarget | undefined => {
        while (cursor < targets.length && (targets[cursor]!.at < at || targets[cursor]!.line !== undefined))
          cursor++
        return targets[cursor]
      }

      breaks.forEach((at, cut) => {
        const ending = metricsOf(sheetOf(cut).section)
        const opening = metricsOf(sheetOf(cut + 1).section)
        const blanks = plan.sheets.slice((contentSheets[cut] ?? 0) + 1, contentSheets[cut + 1] ?? 0)
        const internal = internalAt.get(at)
        const position = internal?.line?.resolve() ?? null
        // Linha não achada (DOM trocado no meio da medida): cede ao bloco seguinte.
        const target =
          internal !== undefined && (internal.line === undefined || position !== null)
            ? internal
            : blockTargetFrom(at)
        // `hangingBottom` cabe na margem de baixo.
        const span = at - previous + liftsBetween(previous, at)
        const hung = Math.min(Math.max(span - ending.contentPx, 0), hangingAt(blocks, at))
        const used = span - hung
        const notesHeight = plan.noteHeights[cut] ?? 0
        sheetStarts.push(previous)
        footnoteArea(cut, ending, used)
        sheetHeights.push(Math.max(ending.heightPx, used + notesHeight + ending.topPx + ending.bottomPx))
        let skipped = 0
        for (const blank of blanks) {
          const height = metricsOf(blank.section).heightPx
          sheetHeights.push(height)
          skipped += height + SHEET_GUTTER_PX
        }
        const shift =
          Math.max(ending.contentPx - used, notesHeight, 0) -
          hung +
          ending.bottomPx +
          SHEET_GUTTER_PX +
          skipped +
          opening.topPx
        if (target?.line !== undefined && position !== null) {
          // O papel recorta o parágrafo no mesmo caractere.
          pageStarts.push({ ...target.start, offset: position - target.line.block - 1 })
          lineGaps.set(position, shift)
        } else if (target !== undefined) {
          // O cabeçalho repetido mora no vão; a conta de fluxo desconta os dois.
          const header = target.header?.()
          const extra = header !== undefined && header.height < opening.contentPx / 2 ? header.height : 0
          if (header !== undefined && extra > 0) headers.push(header)
          pageStarts.push(extra > 0 ? { ...target.start, repeatHeader: true } : target.start)
          for (const node of target.nodes) {
            gaps.set(node.position, shift + extra)
            written.set(node.position, shift + extra + node.natural)
          }
          previous = at - extra
          return
        } else {
          // Uma quebra explícita final ainda abre uma folha vazia.
          pageStarts.push({ blockIndex: blocks.length })
        }
        previous = at
      })

      const last = metricsOf(sheetOf(breaks.length).section)
      const bottom = endnoteBlocks.length > 0 ? endnoteTop : textBottom
      const lastSpan = Math.max(bottom - previous, 0) + liftsBetween(previous, bottom + 1)
      const lastHung = Math.min(Math.max(lastSpan - last.contentPx, 0), hangingAt(blocks, bottom))
      const lastNotes = plan.noteHeights[breaks.length] ?? 0
      sheetStarts.push(previous)
      footnoteArea(breaks.length, last, lastSpan - lastHung)
      sheetHeights.push(Math.max(last.heightPx, lastSpan - lastHung + lastNotes + last.topPx + last.bottomPx))

      endnoteBlocks.forEach((block, position) => {
        const note = endnotes[position]!
        sheetStarts.forEach((start, content) => {
          const end = breaks[content] ?? Number.POSITIVE_INFINITY
          const visible = [block.top, ...block.breakpoints]
            .map((at, line) => ({ at, line }))
            .filter(({ at }) => at >= start - 0.5 && at < end - 0.5)
          if (visible.length === 0) return
          const fromLine = visible[0]!.line
          const toLine = visible.at(-1)!.line + 1
          const metrics = metricsOf(sheetOf(content).section)
          const sheet = contentSheets[content] ?? content
          const item: NoteAreaItem = {
            key: note.id,
            index: note.index,
            fromLine,
            toLine,
            clipTopPx: noteLineTop(note, fromLine),
            heightPx: noteSpan(note, fromLine, toLine),
          }
          const area = noteAreas.find(
            (candidate) => candidate.sheet === sheet && candidate.kind === 'endnote',
          )
          if (area !== undefined) {
            noteAreas[noteAreas.indexOf(area)] = { ...area, items: [...area.items, item] }
            return
          }
          // A folha que só continua as notas não repete o separador.
          const opens = position === 0 && fromLine === 0
          noteAreas.push({
            sheet,
            kind: 'endnote',
            topPx: metrics.topPx + (visible[0]!.at - start) - (opens ? NOTE_SEPARATOR_PX : 0),
            leftPx: metrics.leftPx,
            widthPx: metrics.widthPx - metrics.leftPx - metrics.rightPx,
            separator: opens ? 'normal' : null,
            items: [item],
          })
        })
      })
      // A seção par ou ímpar só pede a folha em branco antes de começar.
      const sheets = plan.sheets.slice(0, sheetHeights.length)
      const sheetWidths = sheets.map((sheet) => metricsOf(sheet.section).widthPx)
      const sheetTops: number[] = []
      let stackHeightPx = 0
      for (const height of sheetHeights) {
        sheetTops.push(stackHeightPx)
        stackHeightPx += height + SHEET_GUTTER_PX
      }

      // O desvio vertical das colunas entra no vão; o lateral é uma translação.
      const columnShifts = new Map<number, number>()
      const blockTargets = new Map<number, CutTarget>()
      for (const cut of targets) {
        if (
          cut.line === undefined &&
          cut.start.childIndex === undefined &&
          !blockTargets.has(cut.start.blockIndex)
        )
          blockTargets.set(cut.start.blockIndex, cut)
      }
      const columnMoves: ColumnMove[] = []
      for (const [index, placement] of plan.placements) {
        const node = blockTargets.get(index)?.nodes[0]
        if (node === undefined) continue
        const metrics = metricsOf(blocks[index]?.section ?? 0)
        const dx = placement.column * metrics.columnStepPx
        if (dx !== 0) columnShifts.set(node.position, dx)
        if (placement.lift !== 0) {
          gaps.set(node.position, (gaps.get(node.position) ?? 0) + placement.lift)
          const distance = (written.get(node.position) ?? node.natural) + placement.lift
          written.set(node.position, collapsed(distance, node.collapse ?? 0))
        }
        columnMoves.push({
          blockIndex: index,
          dx,
          narrowerPx:
            metrics.columns > 1
              ? metrics.widthPx - metrics.leftPx - metrics.rightPx - metrics.columnWidthPx
              : 0,
          lift: placement.lift,
          natural: node.natural,
          collapse: node.collapse ?? 0,
        })
      }

      const columnLines: ColumnLine[] = []
      for (const region of plan.regions) {
        const metrics = metricsOf(region.section)
        if (!metrics.separator) continue
        const sheet = contentSheets[region.sheet] ?? region.sheet
        for (let column = 1; column < region.columns; column++) {
          columnLines.push({
            sheet,
            leftPx:
              metrics.leftPx +
              column * metrics.columnStepPx -
              (metrics.columnStepPx - metrics.columnWidthPx) / 2,
            topPx: metrics.topPx + region.top,
            heightPx: region.height,
          })
        }
      }

      // No modo de leitura o mapa do aplicado esvazia junto, senão a medida
      // seguinte descontaria um empurrão que não existe.
      const target = paginated ? gaps : EMPTY_GAPS
      const targetWritten = paginated ? written : EMPTY_GAPS
      const targetLines = paginated ? lineGaps : EMPTY_GAPS
      const targetHeaders = paginated ? headers : []
      const targetColumns = paginated ? columnShifts : EMPTY_GAPS
      const headersKey = JSON.stringify(targetHeaders)

      if (
        !sameGaps(lastColumns.current, targetColumns) ||
        !sameGaps(applied.current, target) ||
        !sameGaps(lastWritten.current, targetWritten) ||
        !sameGaps(lastLines.current, targetLines) ||
        lastHeaders.current !== headersKey
      ) {
        lastHeaders.current = headersKey
        applied.current = target
        lastWritten.current = targetWritten
        lastLines.current = targetLines
        lastColumns.current = targetColumns
        applyPageGaps(editor.view, targetWritten, target, {
          lines: targetLines,
          headers: targetHeaders,
          columns: targetColumns,
        })
      }

      setLayout({
        pages: sheetHeights.length,
        stackHeightPx: stackHeightPx - SHEET_GUTTER_PX,
        sheetTops,
        sheetHeights,
        pageStarts,
        anchors: anchorsFor(
          blocks,
          breaks,
          (content) => ({
            sheet: contentSheets[content] ?? content,
            marginTopPx: metricsOf(sheetOf(content).section).topPx,
          }),
          (index) => plan.placements.get(index)?.lift ?? 0,
        ),
        sheets,
        sheetWidths,
        stackWidthPx: Math.max(0, ...sheetWidths),
        contentSheets,
        columnMoves,
        columnLines,
        noteAreas,
      })
    }

    // Uma medida por quadro: digitar depressa dispararia dezenas por segundo.
    let scheduled = 0
    const schedule = (): void => {
      if (scheduled !== 0) return
      scheduled = requestAnimationFrame(() => {
        scheduled = 0
        measure()
      })
    }

    schedule()

    const observer = new ResizeObserver(schedule)
    observer.observe(element)
    return () => {
      observer.disconnect()
      if (scheduled !== 0) cancelAnimationFrame(scheduled)
    }
  }, [editor, sections, declared, revision, bandsKey, paginated, styles])

  return layout
}

/** A linha vazia da captura e o vão até o bloco que abre a folha seguinte. */
function hangingAt(blocks: readonly MeasuredBlock[], at: number): number {
  const ending = blocks.filter((block) => block.height > 0 && block.top + block.height <= at + 0.5).at(-1)
  if (ending === undefined || (ending.hangingBottom ?? 0) <= 0) return 0
  return ending.hangingBottom! + Math.max(at - (ending.top + ending.height), 0)
}

function sameGaps(left: ReadonlyMap<number, number>, right: ReadonlyMap<number, number>): boolean {
  if (left.size !== right.size) return false
  for (const [index, gap] of left) {
    // Um pixel não vale uma transação: a medida oscila, e o documento tremeria.
    if (!right.has(index) || Math.abs(right.get(index)! - gap) > 0.5) return false
  }
  return true
}

/** Os cortes são fronteiras crescentes: uma varredura só. */
function anchorsFor(
  blocks: readonly MeasuredBlock[],
  breaks: readonly number[],
  sheetOf: (content: number) => { readonly sheet: number; readonly marginTopPx: number },
  liftOf: (index: number) => number = () => 0,
): BlockAnchor[] {
  const anchors: BlockAnchor[] = []
  let page = 0
  let drift = 0

  blocks.forEach((block, index) => {
    while (page < breaks.length && block.top >= breaks[page]!) {
      page += 1
      drift = 0
    }
    drift += liftOf(index)
    const start = page === 0 ? 0 : breaks[page - 1]!
    const { sheet, marginTopPx } = sheetOf(page)
    anchors.push({ pageIndex: sheet, topPx: marginTopPx + (block.top - start) + drift })
  })

  return anchors
}

/** Soma as origens dos offsetParents: uma linha mede a partir da tabela. */
function offsetTopOf(node: HTMLElement): number {
  let top = 0
  let current: HTMLElement | null = node
  while (current !== null) {
    top += current.offsetTop
    current = current.offsetParent instanceof HTMLElement ? current.offsetParent : null
  }
  return top
}

/** A decoração acompanha edições no modelo; a medida lê o empurrão já mapeado. */
function shiftOf(node: HTMLElement): number {
  return Number(node.dataset.pageShift ?? 0)
}

/**
 * Uma cópia das linhas de cabeçalho sobre o vão da linha que abre a folha,
 * presa à primeira célula dela por margens negativas.
 */
function repeatedHeader(
  editor: Editor,
  table: HTMLTableElement,
  headerRows: readonly HTMLTableRowElement[],
  cell: HTMLTableCellElement,
  height: number,
): RepeatedHeader {
  const tableBox = table.getBoundingClientRect()
  const cellBox = cell.getBoundingClientRect()
  const scale = table.offsetWidth > 0 && tableBox.width > 0 ? tableBox.width / table.offsetWidth : 1
  const style = getComputedStyle(cell)
  const left = (cellBox.left - tableBox.left) / scale + cell.clientLeft + parseFloat(style.paddingLeft)
  const natural = parseFloat(style.paddingTop) - shiftOf(cell)
  const colgroup = table.querySelector(':scope > colgroup')?.outerHTML ?? ''
  const html =
    `<table class="${table.className}" style="width:${table.offsetWidth}px;margin:0">${colgroup}` +
    `<tbody>${headerRows.map(cleanHeaderRow).join('')}</tbody></table>`
  return {
    position: editor.view.posAtDOM(cell, 0),
    html,
    height,
    offsetTop: height + Math.max(natural, 0),
    offsetLeft: left,
  }
}

/** Sem seleção, busca, alça e vãos: copiados, apareceriam em cada folha. */
function cleanHeaderRow(row: HTMLTableRowElement): string {
  const copy = row.cloneNode(true) as HTMLTableRowElement
  for (const transient of copy.querySelectorAll(
    '.column-resize-handle, .page-line-gap, .page-repeated-header',
  )) {
    transient.remove()
  }
  for (const element of [copy, ...copy.querySelectorAll<HTMLElement>('*')]) {
    element.classList.remove('selectedCell', 'search-hit', 'search-hit--current', 'ProseMirror-selectednode')
    if (element.hasAttribute('data-page-start')) {
      element.style.removeProperty('padding-top')
      element.style.removeProperty('margin-top')
      element.removeAttribute('data-page-start')
      element.removeAttribute('data-page-shift')
    }
  }
  return copy.outerHTML
}

interface ReferenceRows {
  readonly lineStarts: readonly number[] | null
  readonly children: readonly HTMLElement[]
  readonly childTops: readonly number[]
}

/** Até onde a folha precisa ir para levar a referência. */
function referenceBottom(
  block: HTMLElement,
  reference: HTMLElement,
  top: number,
  height: number,
  { lineStarts, children, childTops }: ReferenceRows,
): number {
  if (lineStarts !== null) {
    const box = block.getBoundingClientRect()
    const scale = block.offsetHeight > 0 && box.height > 0 ? box.height / block.offsetHeight : 1
    // O pé do sobrescrito, e não o topo: ele sobe acima da linha dele.
    const y = (reference.getBoundingClientRect().bottom - box.top) / scale - 1
    const next = lineStarts.find((start) => start > y)
    return top + (next ?? height)
  }
  const index = children.findIndex((child) => child.contains(reference))
  if (index >= 0 && childTops[index + 1] !== undefined) return childTops[index + 1]!
  return top + height
}

/** A altura do corpo da nota e o topo de cada linha dele, na largura em que está. */
function measureNote(body: NoteBody): { height: number; lines: number[] } {
  const element = body.body
  const height = element.offsetHeight
  if (height <= 0) return { height: 0, lines: [0] }
  const box = element.getBoundingClientRect()
  const scale = box.height > 0 ? box.height / height : 1
  const lines: number[] = []
  for (const child of Array.from(element.children)) {
    if (!(child instanceof HTMLElement)) continue
    const top = (child.getBoundingClientRect().top - box.top) / scale
    const measured = /^(P|H[1-6])$/.test(child.tagName) ? measureLines(body.view, child) : null
    if (measured === null || measured.starts.length === 0) lines.push(top)
    else for (const start of measured.starts) lines.push(top + start)
  }
  const sorted = [...new Set(lines.map((line) => Math.round(line * 100) / 100))]
    .filter((line) => line >= 0 && line < height)
    .sort((left, right) => left - right)
  // A primeira linha começa no topo do corpo: a margem de cima vai com ela.
  sorted[0] = 0
  return { height, lines: sorted }
}
