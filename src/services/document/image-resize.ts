/** Em pixels do CSS, os mesmos dos atributos do nó; o gravador converte em EMU (`ImageWriter.cs`). */

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

/** Abaixo disso as oito alças se sobrepõem e não há como crescer de volta. */
export const MIN_IMAGE_PX = 24

export interface ResizeRequest {
  readonly handle: ResizeHandle
  /** O tamanho quando o gesto começou, e não o do quadro anterior: o arrasto é medido do início. */
  readonly start: ImageSize
  readonly deltaX: number
  readonly deltaY: number
  /** Nos cantos, travada por padrão e o `Shift` solta, como no Word; nas bordas, nunca. */
  readonly keepProportion: boolean
  /** Largura da coluna de texto. A imagem não passa dela, como no Word. */
  readonly maxWidth: number
}

/** Inteiro: um atributo com `342.7188` mudaria a cada quadro por ruído e a paginação remediria. */
export function resizedImage(request: ResizeRequest): ImageSize {
  const { handle, start, deltaX, deltaY, maxWidth } = request

  const horizontal = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0
  const vertical = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0

  const ratio = start.height > 0 && start.width > 0 ? start.height / start.width : 0
  const locked = request.keepProportion && isCornerHandle(handle) && ratio > 0

  let width = start.width + horizontal * deltaX
  let height = start.height + vertical * deltaY

  // Na alça travada manda a largura, o eixo que a coluna limita.
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

/** Quando a coluna encolhe; senão o gravador encolheria a imagem sozinho e a tela mentiria. */
export function fittedImage(size: ImageSize, maxWidth: number): ImageSize {
  if (size.width <= maxWidth || size.width <= 0) return size

  const ceiling = Math.max(MIN_IMAGE_PX, maxWidth)
  return {
    width: Math.round(ceiling),
    height: Math.max(1, Math.round((size.height * ceiling) / size.width)),
  }
}
