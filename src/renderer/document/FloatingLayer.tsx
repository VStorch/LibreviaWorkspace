import { useEffect, useRef } from 'react'
import {
  DOMParser as ProseMirrorParser,
  DOMSerializer,
  Fragment,
  Node as ProseMirrorNode,
  type Schema,
} from '@tiptap/pm/model'
import { mmToPx, type DocumentNode, type PageSetup } from '@services/document/model.js'
import { frameOf, placeFloating, type FloatingObject } from '@services/document/floating.js'

/**
 * Fora do `contenteditable`: no Word os objetos ancorados não estão no fluxo, e
 * dentro dele ocupariam altura e seriam apagáveis. Duas camadas, como o
 * `behindDoc` do OOXML. A caixa de texto se edita no lugar: o título da capa é
 * feito disso.
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
  /** A da faixa fica acima da coluna de texto, que cobre a folha inteira e apanharia o clique. */
  variant?: 'body' | 'band'
  onEdit?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  /** A faixa mora na configuração de página: volta pelo endereço da caixa. */
  onEditBand?: ((bid: string, content: DocumentNode[]) => void) | undefined
}): React.JSX.Element | null {
  const visible = objects.filter((item) => item.object.behind === behind)
  if (visible.length === 0) return null

  // `aria-hidden` só enquanto nada ali é editável.
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
    // Em torno do centro, como o Word gira.
    ...(box.rotation === 0 ? {} : { transform: `rotate(${box.rotation}deg)` }),
    // Por último, para não disputar com a posição.
    ...frameOf(placed.object),
  }

  if (placed.object.kind === 'image') {
    return <img className="paper-float" style={style} src={placed.object.src} alt="" draggable={false} />
  }

  if (placed.object.kind === 'rule') {
    return <div className="paper-float paper-float--rule" style={style} />
  }

  // A caixa que traz numeração perdeu o endereço ao trocar o marcador, e não é editável.
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

/** Serializado para DOM, e não por `innerHTML`: o conteúdo vem do documento. */
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

    // Com o foco, quem manda no DOM é o navegador: redesenhar levaria o cursor.
    if (element.contains(document.activeElement)) return

    element.replaceChildren()
    try {
      const nodes = content.map((node) => ProseMirrorNode.fromJSON(schema, node))
      const serializer = DOMSerializer.fromSchema(schema)
      element.appendChild(serializer.serializeFragment(Fragment.fromArray(nodes)))
    } catch {
      // A caixa que o schema não reconhece fica vazia, sem derrubar a página.
    }
  }, [content, schema])

  // No `blur`: o atributo redesenha a folha, e o cursor se perderia.
  const commit = (): void => {
    const element = host.current
    if (element === null || onEdit === undefined) return

    try {
      const parsed = ProseMirrorParser.fromSchema(schema).parse(element)
      const blocks = (parsed.toJSON() as { content?: DocumentNode[] }).content ?? []
      onEdit(blocks)
    } catch {
      // Nem sobrescreve o que estava lá com conteúdo vazio.
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
