import { pageDimensionsMm, type DocumentNode, type PageSetup } from './model.js'

/**
 * Um objeto que não está no fluxo do texto.
 *
 * Espelha `FloatDto` do sidecar. Medidas em milímetros porque são desenhadas em
 * dois lugares com resoluções diferentes — a tela e o papel — e converter uma
 * vez só na origem deixaria um dos dois arredondando de volta.
 */
export interface FloatingObject {
  readonly kind: 'image' | 'text'
  readonly src?: string | undefined
  readonly content?: DocumentNode[] | undefined
  readonly widthMm: number
  readonly heightMm: number
  /** Graus, sentido horário. */
  readonly rotation: number
  readonly hFrom: string
  readonly hOffsetMm?: number | undefined
  readonly hAlign?: string | undefined
  readonly vFrom: string
  readonly vOffsetMm?: number | undefined
  readonly vAlign?: string | undefined
  readonly behind: boolean
  readonly wrap: string
}

/** A caixa já resolvida, em milímetros da borda da folha. */
export interface FloatingBox {
  readonly leftMm: number
  readonly topMm: number
  readonly widthMm: number
  readonly heightMm: number
  readonly rotation: number
  readonly behind: boolean
}

/**
 * Onde o objeto cai na folha.
 *
 * O OOXML dá a posição em relação a uma de várias origens por eixo — a margem, a
 * coluna, a página, o parágrafo — e a origem vertical mais comum é o parágrafo,
 * que só tem posição depois de paginar. Por isso a altura do parágrafo âncora
 * **dentro da folha** entra como parâmetro: é a única parte que este módulo não
 * consegue saber sozinho.
 *
 * A rotação sai como está, sem mexer nas medidas. O Word posiciona a caixa sem
 * girar e depois a gira em torno do centro, que é o que `transform: rotate()`
 * faz — girar as medidas aqui deslocaria o objeto por metade da diferença entre
 * largura e altura.
 */
export function placeFloating(object: FloatingObject, page: PageSetup, anchorTopMm: number): FloatingBox {
  const { width, height } = pageDimensionsMm(page)
  const columnLeft = page.margins.left
  const columnWidth = width - page.margins.left - page.margins.right

  const leftMm = (() => {
    // Alinhamento manda sobre deslocamento: o OOXML traz um ou outro, nunca os
    // dois, e quando há alinhamento o deslocamento não existe.
    if (object.hAlign !== undefined) {
      const box = referenceH(object.hFrom, page, width, columnLeft, columnWidth)
      if (object.hAlign === 'center') return box.start + (box.size - object.widthMm) / 2
      if (object.hAlign === 'right') return box.start + box.size - object.widthMm
      return box.start
    }

    const box = referenceH(object.hFrom, page, width, columnLeft, columnWidth)
    return box.start + (object.hOffsetMm ?? 0)
  })()

  const topMm = (() => {
    const offset = object.vOffsetMm ?? 0
    switch (object.vFrom) {
      case 'page':
        return offset
      case 'topMargin':
        return offset
      case 'bottomMargin':
        return height - page.margins.bottom + offset
      case 'margin':
        return page.margins.top + offset
      // `paragraph` e `line` são a mesma coisa para nós: a linha exata dentro do
      // parágrafo exigiria medir cada linha, e a diferença é de uma entrelinha.
      default:
        return anchorTopMm + offset
    }
  })()

  return {
    leftMm,
    topMm,
    widthMm: object.widthMm,
    heightMm: object.heightMm,
    rotation: object.rotation,
    behind: object.behind,
  }
}

/** A faixa horizontal a que o deslocamento se refere. */
function referenceH(
  from: string,
  page: PageSetup,
  width: number,
  columnLeft: number,
  columnWidth: number,
): { start: number; size: number } {
  switch (from) {
    case 'page':
      return { start: 0, size: width }
    case 'leftMargin':
      return { start: 0, size: page.margins.left }
    case 'rightMargin':
      return { start: width - page.margins.right, size: page.margins.right }
    case 'insideMargin':
      return { start: 0, size: page.margins.left }
    case 'outsideMargin':
      return { start: width - page.margins.right, size: page.margins.right }
    // `margin`, `column` e `character` caem na coluna de texto. Numa página de
    // coluna única — todas as que o leitor produz hoje — as três coincidem.
    default:
      return { start: columnLeft, size: columnWidth }
  }
}

/** Os objetos que um bloco carrega, ou lista vazia. */
export function floatsOf(attrs: Record<string, unknown> | null | undefined): FloatingObject[] {
  const raw = attrs?.['floats']
  return Array.isArray(raw) ? (raw as FloatingObject[]) : []
}
