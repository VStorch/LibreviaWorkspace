import { useEffect, useRef } from 'react'
import {
  DOMParser as ProseMirrorParser,
  DOMSerializer,
  Fragment,
  Node as ProseMirrorNode,
  type Schema,
} from '@tiptap/pm/model'
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
 *
 * Fora do fluxo não quer dizer fora do alcance: a caixa de texto é editável no
 * lugar em que está. A capa do modelo de manual é feita disso — título e
 * subtítulo não estão no fluxo, e sem isto não haveria como escrevê-los.
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
   * De que camada esta é: a do corpo ou a da faixa.
   *
   * A da faixa fica **acima da coluna de texto**. Não é preferência de desenho:
   * a coluna é um retângulo que cobre a folha inteira, margens incluídas, e
   * apanhava o clique destinado à caixa do cabeçalho. Como a faixa mora na
   * margem e o corpo não entra ali, subi-la não esconde nada.
   */
  variant?: 'body' | 'band' 
  onEdit?: ((source: FloatSource, content: DocumentNode[]) => void) | undefined
  /**
   * O texto de uma caixa da faixa mudou.
   *
   * Caminho próprio porque a faixa não mora no documento do editor: ela é a
   * parte OOXML preservada, e vive na configuração de página. O objeto do corpo
   * volta pelo bloco que o ancora; o da faixa, pelo endereço da caixa.
   */
  onEditBand?: ((bid: string, content: DocumentNode[]) => void) | undefined
}): React.JSX.Element | null {
  const visible = objects.filter((item) => item.object.behind === behind)
  if (visible.length === 0) return null

  // `aria-hidden` só enquanto nada ali dentro é editável: esconder do leitor de
  // tela um campo em que se digita seria pior do que a duplicação que ele
  // evitava.
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

/** De onde o objeto saiu, para o texto digitado saber onde voltar. */
export interface FloatSource {
  /** Posição do bloco âncora no documento. */
  readonly pos: number
  /** Índice do objeto na lista do bloco. */
  readonly index: number
}

/** Um objeto e a altura do parágrafo que o ancora, dentro da folha. */
export interface PlacedFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
  /** Ausente no objeto de faixa: cabeçalho e rodapé não se editam por aqui. */
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
    // Em torno do centro, que é como o Word gira: a caixa é posicionada sem
    // girar e o giro acontece depois.
    ...(box.rotation === 0 ? {} : { transform: `rotate(${box.rotation}deg)` }),
  }

  if (placed.object.kind === 'image') {
    return <img className="paper-float" style={style} src={placed.object.src} alt="" draggable={false} />
  }

  // O filete: uma forma rasa e larga, com contorno e sem conteúdo. É assim que
  // o cabeçalho corporativo desenha a linha que corre sob ele.
  if (placed.object.kind === 'rule') {
    return <div className="paper-float paper-float--rule" style={style} />
  }

  // Duas origens, um só caminho de volta: o objeto do corpo volta pelo bloco que
  // o ancora, o da faixa pelo endereço da caixa. A caixa que traz numeração
  // perdeu o endereço na hora de trocar o marcador, e por isso não é editável.
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

    // Redesenhar por baixo de quem está digitando levaria o cursor embora a
    // cada tecla: enquanto a caixa tem o foco, quem manda no DOM é o navegador.
    if (element.contains(document.activeElement)) return

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

  // O texto sai no `blur`, e não a cada tecla: a alteração do atributo redesenha
  // a folha inteira, e redesenhar debaixo do cursor o perderia.
  const commit = (): void => {
    const element = host.current
    if (element === null || onEdit === undefined) return

    try {
      const parsed = ProseMirrorParser.fromSchema(schema).parse(element)
      const blocks = (parsed.toJSON() as { content?: DocumentNode[] }).content ?? []
      onEdit(blocks)
    } catch {
      // Caixa que o schema não reconhece não derruba a página — e, sobretudo,
      // não sobrescreve o que estava lá com um conteúdo vazio.
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
