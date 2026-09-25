import { Extension, type Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { TableMap } from '@tiptap/pm/tables'
import { Plugin } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  cellBordersFromAttr,
  cellLookPatch,
  cellBordersToCss,
  resolvedColumnWidths,
  tableDraftFrom,
  type CellLook,
  type TableDraft,
} from '@services/document/table-format.js'
import { mmToPx, pxToMm } from '@services/document/model.js'

/**
 * Borda, sombreamento e largura de coluna da tabela.
 *
 * O TableKit dá os comandos de estrutura — inserir, remover, mesclar, dividir,
 * alternar cabeçalho — e nenhum de aparência: célula com fundo cinza e célula com
 * a borda de baixo apagada são as duas coisas que todo documento corporativo tem
 * e que o editor não sabia escrever. Os dois atributos entram aqui, e o gravador
 * os leva a `w:tcBorders` e `w:shd`.
 *
 * Extensão separada de `BlockFormat` pelo mesmo motivo que `ParagraphCommands`: lá
 * mora o atributo do parágrafo, aqui o da célula, e junto seriam duas listas de
 * tipos no mesmo arquivo.
 *
 * ## Uma transação por formulário
 *
 * `applyTableDraft` escreve borda, sombreamento, largura de coluna e linha de
 * cabeçalho numa cadeia só. Quatro transações em fila pediriam quatro `Ctrl+Z`
 * para desfazer um clique em "Aplicar" — e cada uma faria a paginação remedir o
 * documento inteiro.
 */

const CELL_TYPES: readonly string[] = ['tableCell', 'tableHeader']

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    tableLook: {
      /**
       * Borda e sombreamento nas células que a seleção cobre.
       *
       * Com uma função, cada célula recebe o que a função devolve a partir do que
       * ela já tem — é como o diálogo aplica só o que a pessoa mexeu.
       */
      setCellLook: (look: CellLook | ((current: CellLook) => CellLook)) => ReturnType
      /**
       * A grade inteira, em pixels do CSS, e não uma coluna só.
       *
       * Inteira porque é assim que o arquivo a guarda: o `w:tblGrid` é uma lista
       * completa, e o gravador descarta a lista em que falte uma medida — largura
       * parcial não é grade. Quem calcula as que o documento não declara é
       * `resolvedColumnWidths`.
       */
      setColumnWidths: (widths: readonly number[]) => ReturnType
    }
  }
}

export const TableLook = Extension.create({
  name: 'tableLook',

  /**
   * A margem de célula também na tela. O `renderHTML` do atributo só chega ao
   * papel: na tela a tabela é desenhada pelo `TableView` redimensionável, que
   * monta o próprio `<table>` e não aplica os atributos. A decoração de nó cai
   * no embrulho dele, e a variável desce às células por herança.
   */
  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          decorations: (state) => {
            const decorations: Decoration[] = []
            state.doc.descendants((node, pos) => {
              if (node.type.name !== 'table') return true
              const css = cellMarginsCss(node.attrs['cellMargins'])
              if (css !== null) {
                decorations.push(
                  Decoration.node(pos, pos + node.nodeSize, { style: `--cell-margins: ${css}` }),
                )
              }
              return true
            })
            return DecorationSet.create(state.doc, decorations)
          },
        },
      }),
    ]
  },

  addGlobalAttributes() {
    return [
      {
        types: ['table'],
        attributes: {
          /**
           * A margem de célula que o Word usa nesta tabela, em twips — cima,
           * direita, baixo, esquerda —, como o leitor a resolveu (tabela, estilo,
           * padrão do Word). Vira a variável que o `padding` das células lê; a
           * tabela sem ela é a tabela nova, com a margem do modelo do editor.
           */
          cellMargins: {
            default: null,
            parseHTML: () => null,
            renderHTML: (attributes: Record<string, unknown>) => {
              const css = cellMarginsCss(attributes['cellMargins'])
              return css === null ? {} : { style: `--cell-margins: ${css}` }
            },
          },
        },
      },
      {
        types: [...CELL_TYPES],
        attributes: {
          /**
           * Os quatro lados em texto canônico. Ver `table-format.ts` para o
           * porquê de ser texto e não objeto.
           */
          borders: {
            default: null,
            // Sem `parseHTML`: o CSS que a renderização produz não volta a ser o
            // atributo. Colar uma tabela de fora traz as bordas que o navegador
            // desenhar, e inventar o atributo a partir do estilo colado daria à
            // célula uma borda que o documento de origem não declarava.
            parseHTML: () => null,
            renderHTML: (attributes: Record<string, unknown>) => {
              const css = cellBordersToCss(cellBordersFromAttr(attributes['borders']))
              return css === '' ? {} : { style: css }
            },
          },

          shading: {
            default: null,
            parseHTML: () => null,
            renderHTML: (attributes: Record<string, unknown>) => {
              const fill = attributes['shading']
              return typeof fill === 'string' && fill !== '' ? { style: `background-color:${fill}` } : {}
            },
          },
        },
      },
    ]
  },

  addCommands() {
    return {
      setCellLook:
        (look) =>
        ({ state, tr, dispatch }) => {
          let touched = false

          for (const range of state.selection.ranges) {
            state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, pos) => {
              if (!CELL_TYPES.includes(node.type.name)) return true
              const current: CellLook = {
                borders: typeof node.attrs['borders'] === 'string' ? node.attrs['borders'] : null,
                shading: typeof node.attrs['shading'] === 'string' ? node.attrs['shading'] : null,
              }
              const next = typeof look === 'function' ? look(current) : look

              // Célula que não muda não entra na transação: um passo que regrava
              // o mesmo valor ainda conta como edição, e a tabela inteira seria
              // dada como alterada — e regravada — por causa de uma célula.
              if (next.borders !== current.borders) tr.setNodeAttribute(pos, 'borders', next.borders)
              if (next.shading !== current.shading) tr.setNodeAttribute(pos, 'shading', next.shading)
              touched = true
              return false
            })
          }

          if (!touched) return false
          if (dispatch !== undefined) dispatch(tr)
          return true
        },

      setColumnWidths:
        (widths) =>
        ({ state, tr, dispatch }) => {
          const found = tableAt(state.selection.$from)
          if (found === null) return false

          const map = TableMap.get(found.node)
          if (map.width !== widths.length) return false

          // Percorrido pela **grade**, e não pelas células: uma célula mesclada
          // ocupa várias colunas e aparece várias vezes no mapa, e escrever nela
          // uma vez por coluna sobrescreveria a medida anterior.
          const written = new Set<number>()

          for (let column = 0; column < map.width; column += 1) {
            for (let row = 0; row < map.height; row += 1) {
              const relative = map.map[row * map.width + column]
              if (relative === undefined || written.has(relative)) continue
              written.add(relative)

              const cell = found.node.nodeAt(relative)
              if (cell === null) continue

              const rect = map.findCell(relative)
              const own = widths.slice(rect.left, rect.right)
              tr.setNodeMarkup(found.start + relative, undefined, { ...cell.attrs, colwidth: own })
            }
          }

          if (dispatch !== undefined) dispatch(tr)
          return true
        },
    }
  },
})

interface FoundTable {
  readonly node: ProseMirrorNode
  /** Posição do primeiro filho da tabela — a base do mapa. */
  readonly start: number
}

function tableAt($pos: {
  depth: number
  node: (depth: number) => ProseMirrorNode
  start: (depth: number) => number
}): FoundTable | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth)
    if (node.type.spec['tableRole'] === 'table') return { node, start: $pos.start(depth) }
  }

  return null
}

/**
 * Onde o cursor está dentro da tabela, como o diálogo precisa ver.
 *
 * `null` quando o cursor não está em tabela nenhuma — é o que desabilita o item
 * de menu em vez de abrir um diálogo sem assunto.
 */
export interface TablePlacement {
  readonly draft: TableDraft
  /** A grade resolvida, inclusive as colunas que o documento não declara. */
  readonly columnWidths: readonly number[]
  /** Índice da coluna do cursor dentro da grade. */
  readonly column: number
}

export function tablePlacementAt(editor: Editor, contentWidthPx: number): TablePlacement | null {
  const { $from } = editor.state.selection
  const found = tableAt($from)
  if (found === null) return null

  const cell = cellAt($from)
  if (cell === null) return null

  const map = TableMap.get(found.node)
  const rect = map.findCell(cell.pos - found.start)

  // A largura declarada por coluna sai da primeira linha, que é de onde o próprio
  // TableKit a lê para desenhar o `colgroup`.
  const declared: (number | null)[] = []
  const first = found.node.firstChild
  if (first !== null) {
    first.forEach((child) => {
      const widths = child.attrs['colwidth']
      const span = Number(child.attrs['colspan']) || 1
      for (let index = 0; index < span; index += 1) {
        const width = Array.isArray(widths) ? Number(widths[index]) : Number.NaN
        declared.push(Number.isFinite(width) && width > 0 ? width : null)
      }
    })
  }

  const columnWidths = resolvedColumnWidths(declared, contentWidthPx)

  return {
    column: rect.left,
    columnWidths,
    draft: tableDraftFrom({
      borders: cell.node.attrs['borders'],
      shading: cell.node.attrs['shading'],
      // Arredondada em um décimo de milímetro: o campo mostra a medida, não o
      // arredondamento do pixel.
      columnWidthMm: Math.round(pxToMm(columnWidths[rect.left] ?? 0) * 10) / 10,
      headerRow: found.node.firstChild?.firstChild?.type.name === 'tableHeader',
    }),
  }
}

function cellAt($pos: {
  depth: number
  node: (depth: number) => ProseMirrorNode
  before: (depth: number) => number
}): { readonly node: ProseMirrorNode; readonly pos: number } | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth)
    if (CELL_TYPES.includes(node.type.name)) return { node, pos: $pos.before(depth) }
  }

  return null
}

/**
 * O formulário inteiro, numa transação.
 *
 * A linha de cabeçalho vem por último porque `toggleHeaderRow` troca o **tipo**
 * dos nós da primeira linha, e uma troca de tipo invalida as posições que os
 * passos anteriores usaram.
 */
export function applyTableDraft(editor: Editor, draft: TableDraft, contentWidthPx: number): boolean {
  const placement = tablePlacementAt(editor, contentWidthPx)
  if (placement === null) return false

  const chain = editor.chain().focus()

  // Só o que a pessoa mexeu, contado contra o formulário de abertura — ver
  // `cellLookPatch`. Aplicar o rascunho inteiro apagava as bordas que o resumo do
  // formulário não representa.
  const opened = placement.draft
  chain.setCellLook((current) => cellLookPatch(current, opened, draft))

  // A largura também só quando mudou: regravar a grade que ninguém tocou
  // declararia em pixels as colunas que o arquivo media em twips.
  if (draft.columnWidthMm !== null && draft.columnWidthMm !== opened.columnWidthMm) {
    const widths = [...placement.columnWidths]
    widths[placement.column] = Math.max(1, Math.round(mmToPx(draft.columnWidthMm)))
    chain.setColumnWidths(widths)
  }

  if (draft.headerRow !== placement.draft.headerRow) chain.toggleHeaderRow()

  return chain.run()
}

/** `"0 108 0 108"` (twips) em `padding` de CSS; o que não for quatro números some. */
export function cellMarginsCss(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const sides = value.trim().split(/\s+/).map(Number)
  if (sides.length !== 4 || sides.some((side) => !Number.isFinite(side) || side < 0)) return null
  // Twips para pixels de CSS: 1440 por polegada, 96 px por polegada.
  return sides.map((side) => `${Math.round((side / 15) * 100) / 100}px`).join(' ')
}
