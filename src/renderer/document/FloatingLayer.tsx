import { useEffect, useRef } from 'react'
import {
  DOMParser as ProseMirrorParser,
  DOMSerializer,
  Fragment,
  Node as ProseMirrorNode,
  type Schema,
} from '@tiptap/pm/model'
import { type DocumentNode, type PageSetup } from '@services/document/model.js'
import { mmToPx } from '@services/units.js'
import { frameOf, placeFloating, type FloatingObject } from '@services/document/floating.js'

/**
 * Outside the `contenteditable`: in Word anchored objects are not in the flow, and inside it they
 * would take up height and be deletable. Two layers, like OOXML's `behindDoc`. A text box is edited
 * in place: the cover title is made of one.
 */
export function FloatingLayer({
  objects,
  page,
  schema,
  behind,
  variant = 'body',
  onEdit,
  onEditBand,
}: {
  objects: readonly PlacedFloat[]
  page: PageSetup
  schema: Schema
  behind: boolean
  /**
   * The band's sits above the text column, which covers the whole sheet and would catch the click.
   */
  variant?: 'body' | 'band'
  onEdit?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  /** The band lives in the page setup: it goes back through the box address. */
  onEditBand?: ((bid: string, content: DocumentNode[]) => void) | undefined
}): React.JSX.Element | null {
  const visible = objects.filter((item) => item.object.behind === behind)
  if (visible.length === 0) return null

  // `aria-hidden` only while nothing there is editable.
  const editable =
    (onEdit !== undefined && visible.some((item) => item.source !== undefined)) ||
    (onEditBand !== undefined && visible.some((item) => item.object.bid !== undefined))

  return (
    <div
      className={`paper-floats paper-floats--${variant === 'band' ? 'band-' : ''}${behind ? 'behind' : 'front'}`}
      {...(editable ? {} : { 'aria-hidden': true as const })}
    >
      {visible.map((item, index) => (
        <Floating
          key={index}
          placed={item}
          page={page}
          schema={schema}
          onEdit={onEdit}
          onEditBand={onEditBand}
        />
      ))}
    </div>
  )
}

export interface FloatSource {
  readonly pos: number
  readonly index: number
}

export interface PlacedFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
  readonly source?: FloatSource | undefined
}

function Floating({
  placed,
  page,
  schema,
  onEdit,
  onEditBand,
}: {
  placed: PlacedFloat
  page: PageSetup
  schema: Schema
  onEdit?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  onEditBand?: ((bid: string, content: DocumentNode[]) => void) | undefined
}): React.JSX.Element {
  const box = placeFloating(placed.object, page, placed.anchorTopMm)

  const style: React.CSSProperties = {
    left: `${mmToPx(box.leftMm)}px`,
    top: `${mmToPx(box.topMm)}px`,
    width: `${mmToPx(box.widthMm)}px`,
    height: `${mmToPx(box.heightMm)}px`,
    // Around the center, as Word rotates.
    ...(box.rotation === 0 ? {} : { transform: `rotate(${box.rotation}deg)` }),
    // Last, so it does not compete with the position.
    ...frameOf(placed.object),
  }

  if (placed.object.kind === 'image') {
    return <img className="paper-float" style={style} src={placed.object.src} alt="" draggable={false} />
  }

  if (placed.object.kind === 'rule') {
    return <div className="paper-float paper-float--rule" style={style} />
  }

  // A box with numbering lost its address when its bullet changed, and is not editable.
  const source = placed.source
  const bid = placed.object.bid

  const commit =
    onEdit !== undefined && source !== undefined
      ? (content: DocumentNode[]): void => onEdit(source, content)
      : onEditBand !== undefined && bid !== undefined
        ? (content: DocumentNode[]): void => onEditBand(bid, content)
        : undefined

  return (
    <FloatingText
      style={style}
      content={placed.object.content ?? []}
      schema={schema}
      {...(commit === undefined ? {} : { onEdit: commit })}
    />
  )
}

/** Serialized to DOM, not through `innerHTML`: the content comes from the document. */
function FloatingText({
  style,
  content,
  schema,
  onEdit,
}: {
  style: React.CSSProperties
  content: readonly DocumentNode[]
  schema: Schema
  onEdit?: ((content: DocumentNode[]) => void) | undefined
}): React.JSX.Element {
  const host = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const element = host.current
    if (element === null) return

    // With focus, the browser owns the DOM: redrawing would take the cursor away.
    if (element.contains(document.activeElement)) return

    element.replaceChildren()
    try {
      const nodes = content.map((node) => ProseMirrorNode.fromJSON(schema, node))
      const serializer = DOMSerializer.fromSchema(schema)
      element.appendChild(serializer.serializeFragment(Fragment.fromArray(nodes)))
    } catch {
      // A box the schema does not recognize stays empty, without bringing the page down.
    }
  }, [content, schema])

  // On `blur`: the attribute redraws the sheet, and the cursor would be lost.
  const commit = (): void => {
    const element = host.current
    if (element === null || onEdit === undefined) return

    try {
      const parsed = ProseMirrorParser.fromSchema(schema).parse(element)
      const blocks = (parsed.toJSON() as { content?: DocumentNode[] }).content ?? []
      onEdit(blocks)
    } catch {
      // Nor does it overwrite what was there with empty content.
    }
  }

  return (
    <div
      ref={host}
      className={`paper-float paper-float--text page__content${onEdit === undefined ? '' : ' paper-float--edit'}`}
      style={style}
      {...(onEdit === undefined ? {} : { contentEditable: true, suppressContentEditableWarning: true })}
      onBlur={onEdit === undefined ? undefined : commit}
    />
  )
}
