import { linesOf, type Band, type BandCell, type BandPiece } from '@services/document/model.js'

/**
 * Cabeçalho ou rodapé do documento importado, desenhado na margem.
 *
 * **Não é editável, e isso é proposital.** O que volta para o `.docx` é a parte
 * OOXML original, preservada intacta pela gravação cirúrgica; o que está aqui é
 * uma representação para você ver o documento como ele é. Editar esta faixa
 * exigiria regenerar o cabeçalho, e aí o logotipo e o posicionamento se
 * perderiam — exatamente o que a Fase 4 evita.
 *
 * `pointer-events: none` no CSS: clicar aqui não tira o cursor do texto.
 */
export function PageBand({
  band,
  kind,
  pageNumber,
  totalPages,
  insetPx,
  offsetPx,
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
}): React.JSX.Element {
  return (
    <div
      className={`band band--${kind}${band.rule ? ' band--ruled' : ''}`}
      style={{
        left: `${insetPx}px`,
        right: `${insetPx}px`,
        [kind === 'header' ? 'top' : 'bottom']: `${offsetPx}px`,
      }}
      aria-hidden="true"
    >
      {band.rows.length === 0 ? null : (
        <BandGrid band={band} pageNumber={pageNumber} totalPages={totalPages} />
      )}
      <BandCellPieces pieces={band.left} place="left" pageNumber={pageNumber} totalPages={totalPages} />
      <BandCellPieces pieces={band.center} place="center" pageNumber={pageNumber} totalPages={totalPages} />
      <BandCellPieces pieces={band.right} place="right" pageNumber={pageNumber} totalPages={totalPages} />
    </div>
  )
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
function BandGrid({
  band,
  pageNumber,
  totalPages,
}: {
  band: Band
  pageNumber: number
  totalPages: number
}): React.JSX.Element {
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
                    {line.map(renderPiece(pageNumber, totalPages))}
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
  (pageNumber: number, totalPages: number) =>
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
          : piece.text

    return (
      <span
        key={index}
        style={{
          fontWeight: piece.bold ? 700 : undefined,
          fontStyle: piece.italic ? 'italic' : undefined,
          color: piece.color,
          fontSize: piece.fontSize,
          fontFamily: piece.fontFamily,
        }}
      >
        {text}
      </span>
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
  pageNumber,
  totalPages,
}: {
  pieces: readonly BandPiece[]
  place: 'left' | 'center' | 'right'
  pageNumber: number
  totalPages: number
}): React.JSX.Element {
  return (
    <div className={`band__cell band__cell--${place}`}>
      {linesOf(pieces).map((line, index) => (
        <div key={index} className="band__line">
          {line.map(renderPiece(pageNumber, totalPages))}
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
