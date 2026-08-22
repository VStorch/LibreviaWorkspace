import { useEffect, useRef } from 'react'
import { DOMSerializer, Fragment, Node as ProseMirrorNode, type Schema } from '@tiptap/pm/model'
import { mmToPx, type DocumentNode, type PageSetup } from '@services/document/model.js'
import { placeFloating, type FloatingObject } from '@services/document/floating.js'

/**
 * Os objetos ancorados de uma folha.
 *
 * Ficam **fora** do `contenteditable`, na mesma camada das folhas e das faixas.
 * No Word eles não estão no fluxo: não empurram o texto, moram numa posição da
 * página e podem ficar atrás dela. Pô-los dentro do texto editável faria cada um
 * ocupar altura que não ocupa no papel — era assim que a marca vertical da capa
 * virava uma faixa deitada de página inteira — e ainda os deixaria selecionáveis
 * e apagáveis, quando o editor não sabe recriá-los.
 *
 * Duas camadas, e não uma: `behindDoc` do OOXML é decoração de capa e marca
 * d'água, que precisam ficar **debaixo** do texto. O resto fica por cima.
 */
export function FloatingLayer({
  objects,
  page,
  schema,
  behind,
}: {
  objects: readonly PlacedFloat[]
  page: PageSetup
  schema: Schema
  behind: boolean
}): React.JSX.Element | null {
  const visible = objects.filter((item) => item.object.behind === behind)
  if (visible.length === 0) return null

  return (
    <div className={`paper-floats paper-floats--${behind ? 'behind' : 'front'}`} aria-hidden="true">
      {visible.map((item, index) => (
        <Floating key={index} placed={item} page={page} schema={schema} />
      ))}
    </div>
  )
}

/** Um objeto e a altura do parágrafo que o ancora, dentro da folha. */
export interface PlacedFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
}

function Floating({
  placed,
  page,
  schema,
}: {
  placed: PlacedFloat
  page: PageSetup
  schema: Schema
}): React.JSX.Element {
  const box = placeFloating(placed.object, page, placed.anchorTopMm)

  const style: React.CSSProperties = {
    left: `${mmToPx(box.leftMm)}px`,
    top: `${mmToPx(box.topMm)}px`,
    width: `${mmToPx(box.widthMm)}px`,
    height: `${mmToPx(box.heightMm)}px`,
    // Em torno do centro, que é como o Word gira: a caixa é posicionada sem
    // girar e o giro acontece depois.
    ...(box.rotation === 0 ? {} : { transform: `rotate(${box.rotation}deg)` }),
  }

  if (placed.object.kind === 'image') {
    return <img className="paper-float" style={style} src={placed.object.src} alt="" draggable={false} />
  }

  return <FloatingText style={style} content={placed.object.content ?? []} schema={schema} />
}

/**
 * O texto de uma caixa, montado a partir dos nós.
 *
 * Serializado para DOM de verdade e anexado, em vez de virar string de HTML e
 * voltar por `innerHTML`. O conteúdo vem do documento, que é dado não confiável;
 * o serializador constrói a partir do schema e só emite o que ele conhece,
 * enquanto o caminho pela string reabriria a porta do analisador de HTML.
 */
function FloatingText({
  style,
  content,
  schema,
}: {
  style: React.CSSProperties
  content: readonly DocumentNode[]
  schema: Schema
}): React.JSX.Element {
  const host = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const element = host.current
    if (element === null) return

    element.replaceChildren()
    try {
      const nodes = content.map((node) => ProseMirrorNode.fromJSON(schema, node))
      const serializer = DOMSerializer.fromSchema(schema)
      element.appendChild(serializer.serializeFragment(Fragment.fromArray(nodes)))
    } catch {
      // Caixa que o schema não reconhece não derruba a página: ela fica vazia,
      // e o aviso de "formas e caixas de texto" já diz que algo não é
      // reproduzido por inteiro.
    }
  }, [content, schema])

  return <div ref={host} className="paper-float paper-float--text page__content" style={style} />
}
