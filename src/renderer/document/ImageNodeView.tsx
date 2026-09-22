import { useRef, useState } from 'react'
import { NodeViewWrapper, type NodeViewProps } from '@tiptap/react'
import {
  MIN_IMAGE_PX,
  RESIZE_HANDLES,
  type ResizeHandle,
  isCornerHandle,
  resizedImage,
  type ImageSize,
} from '@services/document/image-resize.js'

/** O que cada alça diz ao leitor de tela, e o cursor que ela mostra. */
const HANDLES: Record<ResizeHandle, { readonly label: string; readonly cursor: string }> = {
  nw: { label: 'pelo canto superior esquerdo', cursor: 'nwse-resize' },
  n: { label: 'pela borda de cima', cursor: 'ns-resize' },
  ne: { label: 'pelo canto superior direito', cursor: 'nesw-resize' },
  e: { label: 'pela borda da direita', cursor: 'ew-resize' },
  se: { label: 'pelo canto inferior direito', cursor: 'nwse-resize' },
  s: { label: 'pela borda de baixo', cursor: 'ns-resize' },
  sw: { label: 'pelo canto inferior esquerdo', cursor: 'nesw-resize' },
  w: { label: 'pela borda da esquerda', cursor: 'ew-resize' },
}

/** Passo do teclado: oito pixels, como o Word move objeto com as setas. */
const KEYBOARD_STEP = 8

/**
 * A imagem com alças de redimensionamento.
 *
 * ## Uma transação por gesto
 *
 * É a regra que decide o desenho deste componente. Enquanto o ponteiro anda, o
 * tamanho vive num estado local e só o CSS muda; a transação sai **uma vez**, no
 * `pointerup`. Uma transação por pixel arrastado encheria o histórico — desfazer
 * pediria centenas de `Ctrl+Z` para voltar um arrasto — e faria a paginação
 * remedir a folha em cada quadro, porque é a mudança do documento que a dispara.
 *
 * ## Por que o tamanho vem do atributo, e não do arquivo
 *
 * `width` e `height` são o tamanho que o **documento** pede, que não precisa ser a
 * proporção dos bytes: quem arrasta um canto com a proporção solta estica a
 * imagem, e o Word desenha esticado. O `aspect-ratio` sem `auto` é o que faz o
 * navegador respeitar isso mesmo depois de decodificar a imagem — ver
 * `document-image.ts`, que explica o resto.
 *
 * ## O teto
 *
 * A largura da coluna, medida no elemento que abriga a imagem. Medida, e não
 * lida da configuração de página, porque a mesma imagem pode estar dentro de uma
 * célula de tabela — e ali a coluna é a da célula.
 */
export function ImageNodeView({ node, selected, editor, getPos }: NodeViewProps): React.JSX.Element {
  const frame = useRef<HTMLSpanElement>(null)
  const [dragged, setDragged] = useState<ImageSize | null>(null)

  const attributeSize = sizeOf(node.attrs)
  const size = dragged ?? attributeSize
  const editable = editor.isEditable

  function maxWidth(): number {
    const available = containingBlockOf(frame.current)?.clientWidth ?? 0
    return available > MIN_IMAGE_PX ? available : Number.MAX_SAFE_INTEGER
  }

  /**
   * O tamanho novo, numa transação que só troca os dois atributos.
   *
   * Não pelo `updateAttributes` do NodeView: ele regrava o nó com
   * `setNodeMarkup`, que numa folha como a imagem é um **substituir** — a imagem
   * é apagada e posta de novo. A seleção do nó não sobrevive a isso, e as alças
   * sumiam depois da primeira seta. O `setNodeAttribute` é um passo de atributo:
   * o nó continua o mesmo, a seleção fica onde estava e o desfazer volta o passo.
   */
  function resize(next: ImageSize): void {
    editor.commands.command(({ tr }) => {
      const pos = getPos()
      if (typeof pos !== 'number') return false
      tr.setNodeAttribute(pos, 'width', next.width).setNodeAttribute(pos, 'height', next.height)
      return true
    })
  }

  function startResize(handle: ResizeHandle, event: React.PointerEvent): void {
    if (!editable) return

    // O tamanho de partida é o que está na tela, e não o do atributo: a imagem
    // pode ter chegado sem medida nenhuma, e aí quem a define é o navegador.
    const measured = frame.current?.querySelector('img')?.getBoundingClientRect()
    const start = attributeSize ?? {
      width: Math.round(measured?.width ?? MIN_IMAGE_PX),
      height: Math.round(measured?.height ?? MIN_IMAGE_PX),
    }

    const originX = event.clientX
    const originY = event.clientY
    const ceiling = maxWidth()
    let last = start

    const target = event.currentTarget as HTMLElement
    target.setPointerCapture(event.pointerId)

    const move = (moved: PointerEvent): void => {
      last = resizedImage({
        handle,
        start,
        deltaX: moved.clientX - originX,
        deltaY: moved.clientY - originY,
        // Nos cantos a proporção trava e o `Shift` **solta**: esticar uma captura
        // de tela sem querer é dano que só se percebe no papel.
        keepProportion: !moved.shiftKey,
        maxWidth: ceiling,
      })
      setDragged(last)
    }

    const finish = (): void => {
      target.removeEventListener('pointermove', move)
      target.removeEventListener('pointerup', finish)
      target.removeEventListener('pointercancel', finish)

      // A única transação do gesto. Depois dela o estado local sai de cena, e o
      // que a tela mostra volta a ser o atributo do nó.
      setDragged(null)
      if (last.width !== start.width || last.height !== start.height) {
        resize(last)
      }
    }

    target.addEventListener('pointermove', move)
    target.addEventListener('pointerup', finish)
    target.addEventListener('pointercancel', finish)
    event.preventDefault()
  }

  function nudge(handle: ResizeHandle, event: React.KeyboardEvent): void {
    if (!editable || attributeSize === null) return

    const steps: Record<string, [number, number]> = {
      ArrowLeft: [-KEYBOARD_STEP, 0],
      ArrowRight: [KEYBOARD_STEP, 0],
      ArrowUp: [0, -KEYBOARD_STEP],
      ArrowDown: [0, KEYBOARD_STEP],
    }

    const step = steps[event.key]
    if (step === undefined) return
    event.preventDefault()
    event.stopPropagation()

    const next = resizedImage({
      handle,
      start: attributeSize,
      deltaX: step[0],
      deltaY: step[1],
      keepProportion: !event.shiftKey,
      maxWidth: maxWidth(),
    })
    resize(next)
  }

  const alt = typeof node.attrs['alt'] === 'string' ? (node.attrs['alt'] as string) : ''

  return (
    <NodeViewWrapper
      as="span"
      ref={frame}
      className={dragged === null ? 'image-frame' : 'image-frame image-frame--resizing'}
    >
      <img
        src={typeof node.attrs['src'] === 'string' ? (node.attrs['src'] as string) : ''}
        alt={alt}
        // O tamanho como estilo, e não como atributo: o `height: auto` da folha de
        // estilo do documento venceria o atributo, e a imagem esticada voltaria à
        // proporção do arquivo.
        style={
          size === null
            ? undefined
            : { width: `${size.width}px`, aspectRatio: `${size.width} / ${size.height}` }
        }
        draggable={false}
      />

      {/* As alças só existem para quem pode editar, e só na imagem selecionada:
          oito quadradinhos em volta de toda imagem do documento seriam ruído. */}
      {editable &&
        selected &&
        RESIZE_HANDLES.map((handle) => (
          <button
            key={handle}
            type="button"
            className={`image-frame__grip image-frame__grip--${handle}`}
            style={{ cursor: HANDLES[handle].cursor }}
            contentEditable={false}
            aria-label={`Redimensionar imagem ${HANDLES[handle].label}`}
            title={
              isCornerHandle(handle)
                ? 'Arraste para redimensionar. Shift solta a proporção.'
                : 'Arraste para redimensionar num eixo só.'
            }
            onPointerDown={(event) => startResize(handle, event)}
            onKeyDown={(event) => nudge(handle, event)}
            // Sem isto o `mousedown` tira a seleção do nó e as alças desaparecem
            // antes de o arrasto começar.
            onMouseDown={(event) => event.preventDefault()}
          />
        ))}
    </NodeViewWrapper>
  )
}

function sizeOf(attrs: Record<string, unknown>): ImageSize | null {
  const width = Number(attrs['width'])
  const height = Number(attrs['height'])
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null
  return { width: Math.round(width), height: Math.round(height) }
}

/**
 * O primeiro ancestral que é bloco: é a largura dele que limita a imagem.
 *
 * Não o pai direto. A imagem é conteúdo de linha, e o `ReactRenderer` a embrulha
 * num `span` — cuja largura, para o `clientWidth`, é zero. Medido ali, o teto
 * virava infinito e a alça deixava a imagem crescer para fora da coluna.
 */
function containingBlockOf(element: HTMLElement | null): HTMLElement | null {
  let current = element?.parentElement ?? null
  while (current !== null && getComputedStyle(current).display.startsWith('inline')) {
    current = current.parentElement
  }
  return current
}
