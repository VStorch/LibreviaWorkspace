import type { EditorView } from '@tiptap/pm/view'

/**
 * As linhas de um parágrafo, como o navegador as quebrou.
 *
 * É o que deixa a folha terminar **no meio** de um parágrafo, como no Word: o
 * paginador só corta onde lhe dizem que dá, e até aqui ninguém dizia nada sobre
 * parágrafos — um parágrafo de meia página que não cabia descia inteiro, e a
 * folha de cima ficava com um buraco que o Word e o PDF do LibreOffice não têm.
 *
 * A medida vem dos retângulos de texto (`Range.getClientRects()`), um por
 * pedaço de linha. O retângulo cobre a área de conteúdo da fonte, e não a caixa
 * de linha, que é maior pela entrelinha; a fronteira entre duas linhas é o
 * ponto médio entre o pé de uma e o topo da outra — exato quando as duas linhas
 * têm a mesma fonte, porque a entrelinha se reparte igualmente acima e abaixo.
 */
export interface ParagraphLines {
  /**
   * Onde cada linha, a partir da segunda, começa: em pixels de CSS, a partir do
   * topo da borda do bloco, **sem** os vãos de página já aplicados dentro dele.
   */
  readonly starts: readonly number[]
  /** Soma dos vãos de linha aplicados dentro do bloco. */
  readonly shift: number
  /** Posição no documento do primeiro caractere da linha que `starts[index]` abre. */
  readonly positionOf: (index: number) => number | null
}

/** O espaçador que a paginação põe entre duas linhas de um parágrafo cortado. */
export const LINE_GAP_CLASS = 'page-line-gap'

interface Piece {
  readonly top: number
  readonly bottom: number
  /** Em coordenadas de tela, para a busca do caractere que abre a linha. */
  readonly clientTop: number
  readonly node: Node
  /** Índice do retângulo dentro do nó de texto: acima de zero, a linha começa no meio dele. */
  readonly rectIndex: number
}

interface Line {
  top: number
  bottom: number
  readonly first: Piece
}

export function measureLines(view: EditorView, element: HTMLElement): ParagraphLines {
  const box = element.getBoundingClientRect()
  // A tela pode estar com zoom (`transform`), e o retângulo vem na escala dela;
  // a paginação conta em pixels de CSS, que é o que `offsetHeight` dá.
  const scale = element.offsetHeight > 0 && box.height > 0 ? box.height / element.offsetHeight : 1

  const pieces: Piece[] = []
  let shift = 0
  const push = (rect: DOMRect, node: Node, rectIndex: number): void => {
    if (rect.height <= 0) return
    pieces.push({
      top: (rect.top - box.top) / scale - shift,
      bottom: (rect.bottom - box.top) / scale - shift,
      clientTop: rect.top,
      node,
      rectIndex,
    })
  }

  const walker = document.createTreeWalker(element, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => {
      if (!(node instanceof HTMLElement)) return NodeFilter.FILTER_ACCEPT
      // O vão de uma passada anterior: não é linha, é o empurrão, e tudo que
      // vem depois dele no DOM está abaixo dele na tela.
      if (node.classList.contains(LINE_GAP_CLASS)) {
        // Escondido pela medida (`usePagination`): não empurra nada.
        if (node.style.display !== 'none') shift += Number(node.dataset.pageShift ?? 0)
        return NodeFilter.FILTER_REJECT
      }
      // O que flutua não está na linha; o separador do ProseMirror não tem altura.
      if (node.classList.contains('ProseMirror-separator')) return NodeFilter.FILTER_REJECT
      const position = getComputedStyle(node).position
      if (position === 'absolute' || position === 'fixed') return NodeFilter.FILTER_REJECT
      // Átomo (imagem, campo, marca de formatação): ocupa a linha como um todo.
      if (node.tagName === 'IMG' || node.tagName === 'BR' || node.contentEditable === 'false') {
        push(node.getBoundingClientRect(), node, 0)
        return NodeFilter.FILTER_REJECT
      }
      return NodeFilter.FILTER_SKIP
    },
  })

  const range = document.createRange()
  for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
    if (!(node instanceof Text) || node.length === 0) continue
    range.selectNodeContents(node)
    Array.from(range.getClientRects()).forEach((rect, index) => push(rect, node, index))
  }

  // Um pedaço abre linha nova quando o meio dele já passou do pé da linha
  // corrente. O meio, e não o topo: com entrelinha apertada as áreas de fonte
  // de linhas vizinhas se sobrepõem, e um sobrescrito ou uma letra maior na
  // mesma linha têm topos diferentes sem mudar de linha.
  const lines: Line[] = []
  for (const piece of pieces) {
    const current = lines.at(-1)
    const center = (piece.top + piece.bottom) / 2
    if (current === undefined || piece.rectIndex > 0 || center > current.bottom) {
      lines.push({ top: piece.top, bottom: piece.bottom, first: piece })
    } else {
      current.top = Math.min(current.top, piece.top)
      current.bottom = Math.max(current.bottom, piece.bottom)
    }
  }

  const starts: number[] = []
  for (let index = 1; index < lines.length; index++) {
    starts.push((lines[index - 1]!.bottom + lines[index]!.top) / 2)
  }

  return {
    starts,
    shift,
    positionOf: (index) => {
      const line = lines[index + 1]
      return line === undefined ? null : positionOfLine(view, line.first, range)
    },
  }
}

/**
 * O primeiro caractere da linha, no documento.
 *
 * Quando a linha começa no meio de um nó de texto, busca binária pelo primeiro
 * caractere cujo retângulo desceu da linha anterior. Só roda para o corte
 * escolhido, e não para toda linha medida: medir é a cada quadro, cortar é raro.
 */
function positionOfLine(view: EditorView, first: Piece, range: Range): number | null {
  try {
    if (!(first.node instanceof Text) || first.rectIndex === 0) {
      const parent = first.node.parentNode
      if (parent === null) return null
      if (first.node instanceof Text) return view.posAtDOM(first.node, 0)
      return view.posAtDOM(parent, Array.prototype.indexOf.call(parent.childNodes, first.node) as number)
    }

    const text = first.node
    // O pedaço anterior do mesmo nó é a linha de cima.
    range.selectNodeContents(text)
    const previousTop = range.getClientRects()[first.rectIndex - 1]?.top ?? first.clientTop - 1
    let low = 0
    let high = text.length - 1
    let found = text.length
    while (low <= high) {
      const middle = (low + high) >> 1
      range.setStart(text, middle)
      range.setEnd(text, middle + 1)
      const rects = range.getClientRects()
      const rect = rects[rects.length - 1]
      // O espaço que sobra no fim da linha fica pendurado nela: a linha nova
      // começa no primeiro caractere que desceu.
      if (rect !== undefined && rect.top > previousTop + 1) {
        found = middle
        high = middle - 1
      } else {
        low = middle + 1
      }
    }
    return view.posAtDOM(text, found)
  } catch {
    return null
  }
}
