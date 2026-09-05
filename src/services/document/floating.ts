import { pageDimensionsMm, type DocumentNode, type PageSetup } from './model.js'
import { bandForPage } from './band.js'

/**
 * Um objeto que não está no fluxo do texto.
 *
 * Espelha `FloatDto` do sidecar. Medidas em milímetros porque são desenhadas em
 * dois lugares com resoluções diferentes — a tela e o papel — e converter uma
 * vez só na origem deixaria um dos dois arredondando de volta.
 */
export interface FloatingObject {
  /**
   * `rule` é o filete: uma forma rasa e larga, com contorno e sem conteúdo, que
   * é como o cabeçalho corporativo desenha a linha sob si.
   */
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
  /**
   * Deslocamento da peça dentro do desenho, somado depois de resolver a âncora.
   *
   * Vem de um grupo de formas: a âncora diz onde o grupo está, e cada peça tem
   * a coordenada dela dentro dele. Somar depois é o que faz a conta funcionar
   * tanto com deslocamento quanto com alinhamento declarado — no segundo caso
   * quem resolve a origem é esta função, e não o arquivo.
   */
  readonly dxMm?: number | undefined
  readonly dyMm?: number | undefined
  /**
   * Onde a caixa deste objeto mora no arquivo, quando ela é editável.
   *
   * Só objetos de faixa o trazem: os do corpo voltam pelo bloco que os ancora.
   * A caixa é regenerada por inteiro quando o texto muda — digitar abre e fecha
   * parágrafos, e endereçar parágrafo a parágrafo quebraria no primeiro Enter.
   */
  readonly bid?: string | undefined
  /**
   * A moldura e o preenchimento da forma, quando dá para reproduzi-los.
   *
   * Cor sólida e traço sólido de uma espessura, que é o caso comum e é o que o
   * CSS desenha. O que não cabe aqui não é desenhado e entra no inventário — e
   * é só disso que o aviso "moldura e preenchimento de formas" passa a falar.
   */
  readonly fill?: string | undefined
  readonly line?: string | undefined
  readonly lineWidthPt?: number | undefined
  readonly dash?: boolean | undefined
}

/**
 * A moldura da forma em CSS, para a tela e o papel desenharem a mesma coisa.
 *
 * Compartilhada de propósito: dois desenhistas com a mesma regra escrita duas
 * vezes é como a tela e o papel divergem.
 *
 * Traço de espessura zero não é traço: o formato o usa para dizer "a mais fina
 * possível" quando há cor, e o leitor já resolve isso — aqui, zero é ausência.
 */
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

/** Um objeto e a altura de onde contar a âncora vertical dele. */
export interface AnchoredFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
}

/**
 * Os objetos ancorados das faixas de uma folha.
 *
 * Repetem em toda folha, porque a faixa repete — e por isso não pertencem a
 * bloco nenhum. A âncora vertical deles se diz relativa ao "parágrafo", e o
 * parágrafo de uma faixa começa na distância que `w:pgMar` declara da borda do
 * papel: no alto para o cabeçalho, contada de baixo para o rodapé.
 *
 * Mora aqui, e não em quem desenha, porque a tela e o papel precisam do mesmo
 * número — é a razão de a conta de posição ser uma só.
 */
export function bandFloatsOf(page: PageSetup, pageNumber: number): AnchoredFloat[] {
  const height = pageDimensionsMm(page).height
  const header = bandForPage(page, pageNumber, 'header')
  const footer = bandForPage(page, pageNumber, 'footer')

  return [
    ...(header?.floats ?? []).map((object) => ({
      object: numbered(object, pageNumber),
      anchorTopMm: page.headerDistanceMm,
    })),
    ...(footer?.floats ?? []).map((object) => ({
      object: numbered(object, pageNumber),
      anchorTopMm: height - page.footerDistanceMm,
    })),
  ]
}

/**
 * Troca os marcadores de numeração pelo número desta folha.
 *
 * A caixa de texto do cabeçalho costuma trazer o campo `PAGE` do Word, e o
 * leitor o entrega como `{n}` — o mesmo marcador que o cabeçalho digitado à mão
 * já usa. Sem esta troca a folha sairia com as chaves escritas nela.
 */
function numbered(object: FloatingObject, pageNumber: number): FloatingObject {
  if (object.kind !== 'text' || object.content === undefined) return object

  const content = object.content.map((node) => replaceMarkers(node, pageNumber))

  // A caixa que traz numeração deixa de ser editável, e é por isso que a troca
  // acontece aqui: o que está na tela é o número desta folha, e devolvê-lo ao
  // arquivo trocaria o campo `PAGE` por um número fixo — o cabeçalho passaria a
  // dizer "3" em todas as folhas, e só se notaria na quarta.
  const marked = JSON.stringify(content) !== JSON.stringify(object.content)

  return { ...object, content, ...(marked ? { bid: undefined } : {}) }
}

function replaceMarkers(node: DocumentNode, pageNumber: number): DocumentNode {
  return {
    ...node,
    ...(typeof node.text === 'string' ? { text: node.text.replaceAll('{n}', String(pageNumber)) } : {}),
    ...(Array.isArray(node.content)
      ? { content: node.content.map((child) => replaceMarkers(child, pageNumber)) }
      : {}),
  }
}
