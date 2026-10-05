import { useRef, useState } from 'react'
import { NodeViewWrapper, type NodeViewProps } from '@tiptap/react'
import type { MessageKey } from '@shared/i18n/index.js'
import { useT } from '../i18n.js'
import { screenScaleOf } from './screen-scale.js'
import {
  MIN_IMAGE_PX,
  RESIZE_HANDLES,
  type ResizeHandle,
  isCornerHandle,
  resizedImage,
  type ImageSize,
} from '@services/document/image-resize.js'

const HANDLE_KEYS: Record<ResizeHandle, MessageKey> = {
  nw: 'document.image.handleNw',
  n: 'document.image.handleN',
  ne: 'document.image.handleNe',
  e: 'document.image.handleE',
  se: 'document.image.handleSe',
  s: 'document.image.handleS',
  sw: 'document.image.handleSw',
  w: 'document.image.handleW',
}

const HANDLE_CURSORS: Record<ResizeHandle, string> = {
  nw: 'nwse-resize',
  n: 'ns-resize',
  ne: 'nesw-resize',
  e: 'ew-resize',
  se: 'nwse-resize',
  s: 'ns-resize',
  sw: 'nesw-resize',
  w: 'ew-resize',
}

/** The way Word moves objects with the arrow keys. */
const KEYBOARD_STEP = 8

/**
 * One transaction per gesture, on `pointerup`: one per pixel would fill the history and remeasure
 * the sheet every frame. The size comes from the attribute, which may stretch the image (see
 * `document-image.ts`). The ceiling is the width measured on the block holding the image, which in
 * a cell is the cell's.
 */
export function ImageNodeView({ node, selected, editor, getPos }: NodeViewProps): React.JSX.Element {
  const frame = useRef<HTMLSpanElement>(null)
  const [dragged, setDragged] = useState<ImageSize | null>(null)

  const attributeSize = sizeOf(node.attrs)
  const size = dragged ?? attributeSize
  const editable = editor.isEditable

  /**
   * `setNodeAttribute`, not `updateAttributes`: the latter would replace the leaf, and the
   * selection and handles would vanish.
   */
  function resize(next: ImageSize): void {
    editor.commands.command(({ tr }) => {
      const pos = getPos()
      if (typeof pos !== 'number') return false
      tr.setNodeAttribute(pos, 'width', next.width).setNodeAttribute(pos, 'height', next.height)
      return true
    })
  }

  function nudge(handle: ResizeHandle, event: React.KeyboardEvent): void {
    if (!editable || attributeSize === null) return
    const step = KEYBOARD_STEPS[event.key]
    if (step === undefined) return
    event.preventDefault()
    event.stopPropagation()

    resize(
      resizedImage({
        handle,
        start: attributeSize,
        deltaX: step[0],
        deltaY: step[1],
        keepProportion: !event.shiftKey,
        maxWidth: maxWidthOf(frame.current),
      }),
    )
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
        // As a style: the stylesheet's `height: auto` would beat the attribute.
        style={
          size === null
            ? undefined
            : { width: `${size.width}px`, aspectRatio: `${size.width} / ${size.height}` }
        }
        draggable={false}
      />
      {/* Only for editors, and only on the selected image. */}
      {editable && selected && (
        <ResizeHandles
          onStart={(handle, event) =>
            dragResize({
              handle,
              event,
              frame: frame.current,
              attributeSize,
              onPreview: setDragged,
              onCommit: resize,
            })
          }
          onNudge={nudge}
        />
      )}
    </NodeViewWrapper>
  )
}

const KEYBOARD_STEPS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowLeft: [-KEYBOARD_STEP, 0],
  ArrowRight: [KEYBOARD_STEP, 0],
  ArrowUp: [0, -KEYBOARD_STEP],
  ArrowDown: [0, KEYBOARD_STEP],
}

function maxWidthOf(frame: HTMLSpanElement | null): number {
  const available = containingBlockOf(frame)?.clientWidth ?? 0
  return available > MIN_IMAGE_PX ? available : Number.MAX_SAFE_INTEGER
}

interface DragResize {
  readonly handle: ResizeHandle
  readonly event: React.PointerEvent
  readonly frame: HTMLSpanElement | null
  readonly attributeSize: ImageSize | null
  /** The size during the gesture; `null` when it ends. */
  readonly onPreview: (size: ImageSize | null) => void
  readonly onCommit: (size: ImageSize) => void
}

function dragResize({ handle, event, frame, attributeSize, onPreview, onCommit }: DragResize): void {
  // The size on screen, since the attribute may be missing; with zoom, measures come back divided
  // by the scale.
  const image = frame?.querySelector('img') ?? null
  const scale = screenScaleOf(image)
  const measured = image?.getBoundingClientRect()
  const start = attributeSize ?? {
    width: Math.round((measured?.width ?? MIN_IMAGE_PX * scale) / scale),
    height: Math.round((measured?.height ?? MIN_IMAGE_PX * scale) / scale),
  }

  const originX = event.clientX
  const originY = event.clientY
  const ceiling = maxWidthOf(frame)
  let last = start

  const target = event.currentTarget as HTMLElement
  target.setPointerCapture(event.pointerId)

  const move = (moved: PointerEvent): void => {
    last = resizedImage({
      handle,
      start,
      deltaX: (moved.clientX - originX) / scale,
      deltaY: (moved.clientY - originY) / scale,
      // On corners the ratio locks and `Shift` releases it, as in Word.
      keepProportion: !moved.shiftKey,
      maxWidth: ceiling,
    })
    onPreview(last)
  }

  const finish = (): void => {
    target.removeEventListener('pointermove', move)
    target.removeEventListener('pointerup', finish)
    target.removeEventListener('pointercancel', finish)

    // The gesture's only transaction.
    onPreview(null)
    if (last.width !== start.width || last.height !== start.height) onCommit(last)
  }

  target.addEventListener('pointermove', move)
  target.addEventListener('pointerup', finish)
  target.addEventListener('pointercancel', finish)
  event.preventDefault()
}

function ResizeHandles({
  onStart,
  onNudge,
}: {
  onStart: (handle: ResizeHandle, event: React.PointerEvent) => void
  onNudge: (handle: ResizeHandle, event: React.KeyboardEvent) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      {RESIZE_HANDLES.map((handle) => (
        <button
          key={handle}
          type="button"
          className={`image-frame__grip image-frame__grip--${handle}`}
          style={{ cursor: HANDLE_CURSORS[handle] }}
          contentEditable={false}
          aria-label={t('document.image.resizeLabel', { handle: t(HANDLE_KEYS[handle]) })}
          title={
            isCornerHandle(handle) ? t('document.image.resizeCornerHint') : t('document.image.resizeEdgeHint')
          }
          onPointerDown={(event) => onStart(handle, event)}
          onKeyDown={(event) => onNudge(handle, event)}
          // Otherwise `mousedown` would drop the selection and the handles before the drag.
          onMouseDown={(event) => event.preventDefault()}
        />
      ))}
    </>
  )
}

function sizeOf(attrs: Record<string, unknown>): ImageSize | null {
  const width = Number(attrs['width'])
  const height = Number(attrs['height'])
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null
  return { width: Math.round(width), height: Math.round(height) }
}

/** The direct parent is the `ReactRenderer` `span`, zero-width for `clientWidth`. */
function containingBlockOf(element: HTMLElement | null): HTMLElement | null {
  let current = element?.parentElement ?? null
  while (current !== null && getComputedStyle(current).display.startsWith('inline')) {
    current = current.parentElement
  }
  return current
}
