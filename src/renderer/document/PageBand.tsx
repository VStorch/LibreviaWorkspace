import type { Band, BandPiece } from '@services/document/model.js'

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
}): React.JSX.Element {
  return (
    <div
      className={`band band--${kind}${band.rule ? ' band--ruled' : ''}`}
      style={{ left: `${insetPx}px`, right: `${insetPx}px` }}
      aria-hidden="true"
    >
      <div className="band__cell band__cell--left">{band.left.map(renderPiece(pageNumber, totalPages))}</div>
      <div className="band__cell band__cell--center">
        {band.center.map(renderPiece(pageNumber, totalPages))}
      </div>
      <div className="band__cell band__cell--right">
        {band.right.map(renderPiece(pageNumber, totalPages))}
      </div>
    </div>
  )
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
        }}
      >
        {text}
      </span>
    )
  }

function sizeOf(piece: BandPiece): React.CSSProperties {
  return piece.width === undefined
    ? {}
    : { width: `${piece.width}px`, height: piece.height === undefined ? 'auto' : `${piece.height}px` }
}
