import { useEffect, useRef } from 'react'
import { linesOf, type Band, type BandCell, type BandPiece } from '@services/document/model.js'

/**
 * Cabeçalho ou rodapé do documento importado, desenhado na margem.
 *
 * **O texto é editável; a moldura não.** Cada peça que trouxe endereço do
 * arquivo recebe o cursor, e o que se digita nela volta para o `w:t` de onde
 * veio. O resto — o logotipo, a tabela, o filete, o campo do número da página —
 * continua sendo desenho de uma parte OOXML que a gravação devolve intacta, e
 * que este componente não saberia gerar de novo.
 *
 * A faixa é uma só no arquivo mesmo aparecendo em todas as folhas: editar na
 * folha 3 muda todas, como no Word.
 *
 * `pointer-events: none` na faixa, e `auto` só nas peças editáveis: clicar na
 * margem em volta não tira o cursor do texto do corpo.
 */
export function PageBand({
  band,
  kind,
  pageNumber,
  totalPages,
  insetPx,
  offsetPx,
  onEdit,
}: {
  band: Band
  kind: 'header' | 'footer'
  /** Em qual folha esta faixa está sendo desenhada. */
  pageNumber: number
  /** Quantas folhas o documento tem agora. */
  totalPages: number
  /**
   * Recuo lateral da faixa. Cabeçalho corporativo costuma ser **mais largo que
   * a coluna de texto** — no corpus real, 177 mm contra 146,5 mm — então usar
   * a margem do texto encolheria o logotipo para dentro. Metade da margem
   * reproduz a proporção do original sem precisar de número mágico.
   */
  insetPx: number
  /**
   * Distância da faixa até a borda do papel, que o arquivo declara em
   * `w:pgMar/@header` e `@footer`.
   *
   * Era um valor fixo de 4 mm no CSS. Enquanto o cabeçalho só precisava caber
   * na margem, a diferença era invisível; agora que a altura dele empurra o
   * corpo para baixo, desenhar num lugar e contar de outro faria a conta e o
   * desenho discordarem — e o corpo desceria demais ou de menos.
   */
  offsetPx: number
  /**
   * O texto de uma peça mudou.
   *
   * Ausente quando o documento está travado — e nesse caso nenhuma peça recebe
   * o cursor, em vez de recebê-lo e descartar o que for digitado.
   */
  onEdit?: ((pid: string, text: string) => void) | undefined
}): React.JSX.Element {
  const parts = { pageNumber, totalPages, onEdit }

  return (
    <div
      className={`band band--${kind}${band.rule ? ' band--ruled' : ''}`}
      style={{
        left: `${insetPx}px`,
        right: `${insetPx}px`,
        [kind === 'header' ? 'top' : 'bottom']: `${offsetPx}px`,
      }}
      // `aria-label` e não `aria-hidden`: esconder uma região que agora tem
      // conteúdo editável a tornaria inalcançável pelo teclado, e o leitor de
      // tela anunciaria um campo que não existe.
      role="group"
      aria-label={kind === 'header' ? 'Cabeçalho' : 'Rodapé'}
    >
      {band.rows.length === 0 ? null : <BandGrid band={band} {...parts} />}
      <BandCellPieces pieces={band.left} place="left" {...parts} />
      <BandCellPieces pieces={band.center} place="center" {...parts} />
      <BandCellPieces pieces={band.right} place="right" {...parts} />
    </div>
  )
}

/** O que cada desenhista de peça precisa saber, junto. */
interface BandParts {
  pageNumber: number
  totalPages: number
  onEdit?: ((pid: string, text: string) => void) | undefined
}

/**
 * A grade do cabeçalho, quando ele é uma tabela.
 *
 * Uma tabela de verdade, e não três colunas: no corpus real o logotipo mora
 * numa célula mesclada por quatro linhas, com o título ao lado e a numeração à
 * direita. Espalhado pelos terços, o mesmo cabeçalho virava uma fileira de
 * palavras que transbordava sobre a primeira linha do texto.
 *
 * As bordas vêm resolvidas do leitor, lado a lado: no OOXML cada uma delas sai
 * de três lugares — a célula, a moldura da tabela, a linha interna — e refazer
 * essa conta aqui e outra vez no papel é como os dois divergem.
 */
function BandGrid({ band, ...parts }: { band: Band } & BandParts): React.JSX.Element {
  return (
    <table className="band__grid">
      <tbody>
        {band.rows.map((row, index) => (
          <tr key={index}>
            {row.cells.map((cell, at) => (
              <td
                key={at}
                colSpan={cell.span === 1 ? undefined : cell.span}
                rowSpan={cell.rowSpan === 1 ? undefined : cell.rowSpan}
                style={cellStyle(cell)}
              >
                {linesOf(cell.pieces).map((line, row) => (
                  <div key={row} className="band__line">
                    {line.map(renderPiece(parts))}
                  </div>
                ))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function cellStyle(cell: BandCell): React.CSSProperties {
  const line = '1px solid currentcolor'
  return {
    ...(cell.width > 0 ? { width: `${(cell.width * 100).toFixed(2)}%` } : {}),
    ...(cell.align === undefined ? {} : { textAlign: cell.align as React.CSSProperties['textAlign'] }),
    borderTop: cell.borders.includes('t') ? line : undefined,
    borderLeft: cell.borders.includes('l') ? line : undefined,
    borderBottom: cell.borders.includes('b') ? line : undefined,
    borderRight: cell.borders.includes('r') ? line : undefined,
  }
}

const renderPiece =
  ({ pageNumber, totalPages, onEdit }: BandParts) =>
  (piece: BandPiece, index: number): React.JSX.Element => {
    if (piece.kind === 'image') {
      return (
        <img
          key={index}
          className="band__image"
          src={piece.src}
          alt=""
          style={sizeOf(piece)}
          draggable={false}
        />
      )
    }

    // O total existia só depois de exportar, e a faixa mostrava um marcador no
    // lugar dele. Agora a tela pagina, então o número é o de verdade — e é ele
    // que a pessoa confere antes de imprimir.
    const text =
      piece.kind === 'pageNumber'
        ? String(pageNumber)
        : piece.kind === 'totalPages'
          ? String(totalPages)
          : (piece.text ?? '')

    const style: React.CSSProperties = {
      fontWeight: piece.bold ? 700 : undefined,
      fontStyle: piece.italic ? 'italic' : undefined,
      color: piece.color,
      fontSize: piece.fontSize,
      fontFamily: piece.fontFamily,
    }

    // Só peça com endereço recebe o cursor: o número da página é calculado a
    // cada abertura e não tem `w:t` onde guardar o que fosse digitado nele.
    if (piece.pid === undefined || onEdit === undefined) {
      return (
        <span key={index} style={style}>
          {text}
        </span>
      )
    }

    return <BandText key={index} pid={piece.pid} text={text} style={style} onEdit={onEdit} />
  }

/**
 * Uma peça de texto da faixa, com o cursor dentro.
 *
 * O texto sai no `blur`, e não a cada tecla: mudar a configuração de página
 * redesenha todas as folhas, e redesenhar por baixo de quem digita levaria o
 * cursor embora. Pela mesma razão o conteúdo só é reescrito quando a peça não
 * tem o foco.
 */
function BandText({
  pid,
  text,
  style,
  onEdit,
}: {
  pid: string
  text: string
  style: React.CSSProperties
  onEdit: (pid: string, text: string) => void
}): React.JSX.Element {
  const host = useRef<HTMLSpanElement>(null)

  useEffect(() => {
    const element = host.current
    if (element === null) return
    if (element === document.activeElement) return
    if (element.textContent === text) return

    element.textContent = text
  }, [text])

  return (
    <span
      ref={host}
      className="band__text"
      style={style}
      contentEditable
      suppressContentEditableWarning
      role="textbox"
      spellCheck={false}
      onBlur={() => {
        onEdit(pid, host.current?.textContent ?? '')
      }}
      // Enter abriria parágrafo dentro da peça, e a faixa não tem onde guardar
      // um: um `w:t` é uma linha só. A tecla passa a fechar a edição.
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault()
          host.current?.blur()
        }
      }}
    />
  )
}

/**
 * Um terço da faixa, com as peças agrupadas em linhas.
 *
 * Cada parágrafo do arquivo é uma linha, e por isso são elementos de verdade e
 * não um `<br>`: dentro de um contêiner flex o `br` não gera caixa nenhuma, e o
 * rodapé de três linhas do modelo de manual continuava saindo como uma frase
 * só, emendada na largura da folha.
 */
function BandCellPieces({
  pieces,
  place,
  ...parts
}: { pieces: readonly BandPiece[]; place: 'left' | 'center' | 'right' } & BandParts): React.JSX.Element {
  return (
    <div className={`band__cell band__cell--${place}`}>
      {linesOf(pieces).map((line, index) => (
        <div key={index} className="band__line">
          {line.map(renderPiece(parts))}
        </div>
      ))}
    </div>
  )
}

function sizeOf(piece: BandPiece): React.CSSProperties {
  return piece.width === undefined
    ? {}
    : { width: `${piece.width}px`, height: piece.height === undefined ? 'auto' : `${piece.height}px` }
}
