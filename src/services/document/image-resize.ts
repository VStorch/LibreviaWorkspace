/**
 * A conta do redimensionamento de imagem por alças.
 *
 * Fora do componente porque é a única parte do gesto que tem resposta certa e
 * errada: o resto é ouvinte de ponteiro. Aqui mora o que decide o tamanho — que
 * lado a alça move, quando a proporção trava, e os dois limites que impedem a
 * imagem de sumir ou de estourar a coluna de texto.
 *
 * A **unidade é o pixel do CSS**, a mesma dos atributos `width` e `height` do nó
 * e a mesma que o gravador converte em EMU (96 px por polegada, 914400 EMU por
 * polegada — ver `ImageWriter.cs`). Nenhuma conversão acontece aqui.
 */

/** Onde a alça está. O nome diz os lados que ela move. */
export const ResizeHandle = {
  NorthWest: 'nw',
  North: 'n',
  NorthEast: 'ne',
  East: 'e',
  SouthEast: 'se',
  South: 's',
  SouthWest: 'sw',
  West: 'w',
} as const
export type ResizeHandle = (typeof ResizeHandle)[keyof typeof ResizeHandle]

export const RESIZE_HANDLES: readonly ResizeHandle[] = Object.values(ResizeHandle)

/** Alça de canto mexe nas duas medidas; a de borda, numa só. */
export function isCornerHandle(handle: ResizeHandle): boolean {
  return handle.length === 2
}

export interface ImageSize {
  readonly width: number
  readonly height: number
}

/**
 * Menor imagem que ainda se vê e ainda se pega pela alça.
 *
 * Abaixo disso as oito alças se sobrepõem e não há como crescer de volta a não
 * ser desfazendo — e desfazer um arrasto que a pessoa não terminou é o tipo de
 * armadilha que faz o recurso ser evitado.
 */
export const MIN_IMAGE_PX = 24

export interface ResizeRequest {
  readonly handle: ResizeHandle
  /** O tamanho quando o gesto começou, e não o do quadro anterior: o arrasto é medido do início. */
  readonly start: ImageSize
  /** Quanto o ponteiro andou desde o início do gesto. */
  readonly deltaX: number
  readonly deltaY: number
  /**
   * Se a proporção original deve ser mantida.
   *
   * Nos cantos vale `true` por padrão e o `Shift` **solta** — é o que o Word faz
   * desde 2013, e é a escolha que erra menos: esticar uma captura de tela sem
   * querer é dano que só se percebe no papel. Nas bordas a proporção nunca trava:
   * a alça do meio existe justamente para mexer numa medida só.
   */
  readonly keepProportion: boolean
  /** Largura da coluna de texto. A imagem não passa dela, como no Word. */
  readonly maxWidth: number
}

/**
 * O tamanho que a imagem fica, já arredondado para pixel inteiro.
 *
 * Inteiro porque é isto que vai para o atributo do nó, e um atributo com
 * `342.7188` faria cada quadro do arrasto produzir um valor diferente do
 * anterior por ruído de ponto flutuante — e a paginação remede a cada mudança.
 */
export function resizedImage(request: ResizeRequest): ImageSize {
  const { handle, start, deltaX, deltaY, maxWidth } = request

  const horizontal = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0
  const vertical = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0

  const ratio = start.height > 0 && start.width > 0 ? start.height / start.width : 0
  const locked = request.keepProportion && isCornerHandle(handle) && ratio > 0

  let width = start.width + horizontal * deltaX
  let height = start.height + vertical * deltaY

  // Na alça travada manda a largura: é o eixo que a coluna de texto limita, e
  // deixar a altura mandar faria a imagem encolher de lado ao bater no teto.
  if (locked) height = width * ratio
  if (horizontal === 0) width = locked ? height / ratio : start.width
  if (vertical === 0) height = locked ? width * ratio : start.height

  const ceiling = Math.max(MIN_IMAGE_PX, maxWidth)
  if (width > ceiling) {
    width = ceiling
    if (locked) height = width * ratio
  }

  if (width < MIN_IMAGE_PX) {
    width = MIN_IMAGE_PX
    if (locked) height = width * ratio
  }

  if (height < MIN_IMAGE_PX) {
    height = MIN_IMAGE_PX
    if (locked) width = Math.min(ceiling, height / ratio)
  }

  return { width: Math.round(width), height: Math.round(height) }
}

/**
 * O tamanho já cabe na coluna? Senão, o que caberia.
 *
 * Usado quando a coluna de texto encolhe — trocar de A4 para Letter, virar a
 * folha, aumentar a margem. A imagem que era do tamanho da coluna passaria a
 * estourá-la, e o gravador a encolheria por conta própria na hora de salvar: o
 * que a tela mostra deixaria de ser o que o arquivo tem.
 */
export function fittedImage(size: ImageSize, maxWidth: number): ImageSize {
  if (size.width <= maxWidth || size.width <= 0) return size

  const ceiling = Math.max(MIN_IMAGE_PX, maxWidth)
  return {
    width: Math.round(ceiling),
    height: Math.max(1, Math.round((size.height * ceiling) / size.width)),
  }
}
