import { pageDimensionsMm, type DocumentNode, type PageSetup } from './model.js'
import { bandForPage, pageLabel } from './band.js'

/** Espelha `FloatDto` do sidecar, em milímetros: tela e papel convertem cada um uma vez. */
export interface FloatingObject {
  /** `rule` é o filete sob o cabeçalho corporativo: forma rasa, com contorno e sem conteúdo. */
  readonly kind: 'image' | 'text' | 'rule'
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
  /** A posição da peça dentro do grupo de formas, somada depois de resolver a âncora. */
  readonly dxMm?: number | undefined
  readonly dyMm?: number | undefined
  /** Só nos objetos de faixa; a caixa é regenerada inteira, porque digitar abre e fecha parágrafos. */
  readonly bid?: string | undefined
  /** Só cor e traço sólidos; o resto não é desenhado e entra no inventário. */
  readonly fill?: string | undefined
  readonly line?: string | undefined
  readonly lineWidthPt?: number | undefined
  readonly dash?: boolean | undefined
}

/** Para a tela e o papel desenharem igual. Traço de espessura zero é ausência. */
export function frameOf(object: FloatingObject): { background?: string; border?: string } {
  const frame: { background?: string; border?: string } = {}

  if (typeof object.fill === 'string' && object.fill.length > 0) frame.background = object.fill

  const width = object.lineWidthPt ?? 0
  if (typeof object.line === 'string' && object.line.length > 0 && width > 0) {
    frame.border = `${width}pt ${object.dash === true ? 'dashed' : 'solid'} ${object.line}`
  }

  return frame
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
 * A origem vertical mais comum é o parágrafo, que só tem posição depois de
 * paginar: por isso `anchorTopMm` entra como parâmetro. A rotação sai como está,
 * porque o Word posiciona a caixa sem girar e gira em torno do centro, como
 * `transform: rotate()`.
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
    leftMm: leftMm + (object.dxMm ?? 0),
    topMm: topMm + (object.dyMm ?? 0),
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
    // `margin`, `column` e `character` coincidem numa página de coluna única.
    default:
      return { start: columnLeft, size: columnWidth }
  }
}

/** Os objetos que um bloco carrega, ou lista vazia. */
export function floatsOf(attrs: Record<string, unknown> | null | undefined): FloatingObject[] {
  const raw = attrs?.['floats']
  return Array.isArray(raw) ? (raw as FloatingObject[]) : []
}

/** Um objeto e a altura de onde contar a âncora vertical dele. */
export interface AnchoredFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
}

/**
 * Repetem em toda folha, como a faixa. O "parágrafo" de uma faixa começa na
 * distância que `w:pgMar` declara: do alto no cabeçalho, de baixo no rodapé.
 */
export function bandFloatsOf(page: PageSetup, pageNumber: number): AnchoredFloat[] {
  const height = pageDimensionsMm(page).height
  const header = bandForPage(page, pageNumber, 'header')
  const footer = bandForPage(page, pageNumber, 'footer')
  const label = pageLabel(page, pageNumber)

  return [
    ...(header?.floats ?? []).map((object) => ({
      object: numbered(object, label),
      anchorTopMm: page.headerDistanceMm,
    })),
    ...(footer?.floats ?? []).map((object) => ({
      object: numbered(object, label),
      anchorTopMm: height - page.footerDistanceMm,
    })),
  ]
}

/** O leitor entrega o campo `PAGE` da caixa como `{n}`; sem a troca a folha sairia com as chaves. */
function numbered(object: FloatingObject, pageNumber: string): FloatingObject {
  if (object.kind !== 'text' || object.content === undefined) return object

  const content = object.content.map((node) => replaceMarkers(node, pageNumber))

  // A caixa com numeração deixa de ser editável: devolver o número desta folha
  // ao arquivo trocaria o campo `PAGE` por um número fixo.
  const marked = JSON.stringify(content) !== JSON.stringify(object.content)

  return { ...object, content, ...(marked ? { bid: undefined } : {}) }
}

function replaceMarkers(node: DocumentNode, pageNumber: string): DocumentNode {
  return {
    ...node,
    ...(typeof node.text === 'string' ? { text: node.text.replaceAll('{n}', pageNumber) } : {}),
    ...(Array.isArray(node.content)
      ? { content: node.content.map((child) => replaceMarkers(child, pageNumber)) }
      : {}),
  }
}
