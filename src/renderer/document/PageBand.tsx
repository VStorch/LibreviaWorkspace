import { useEffect, useRef } from 'react'
import { linesOf, pieceText, type Band, type BandCell, type BandPiece } from '@services/document/band.js'
import { usePreferences } from '../state/preferences.js'
import { useT } from '../i18n.js'

/**
 * **O texto é editável; a moldura não.** Cada peça com endereço do arquivo
 * recebe o cursor, e o que se digita volta ao `w:t` dela; o resto é desenho de
 * uma parte OOXML que volta intacta. `pointer-events: none` na faixa, para o
 * clique na margem não tirar o cursor do corpo.
 */
export function PageBand({
  band,
  kind,
  pageLabel,
  totalPages,
  insetPx,
  offsetPx,
  onEdit,
}: {
  band: Band
  kind: 'header' | 'footer'
  /** No formato de `w:pgNumType`. */
  pageLabel: string
  totalPages: number
  /** Metade da margem: o cabeçalho corporativo é mais largo que a coluna de texto. */
  insetPx: number
  /** `w:pgMar/@header` e `@footer`: desenhar num lugar e contar de outro faria o corpo descer errado. */
  offsetPx: number
  /** Ausente quando o documento está travado: nenhuma peça recebe o cursor. */
  onEdit?: ((pid: string, text: string) => void) | undefined
}): React.JSX.Element {
  const t = useT()
  const parts = { pageLabel, totalPages, onEdit }

  return (
    <div
      className={`band band--${kind}${band.rule ? ' band--ruled' : ''}`}
      style={{
        left: `${insetPx}px`,
        right: `${insetPx}px`,
        [kind === 'header' ? 'top' : 'bottom']: `${offsetPx}px`,
      }}
      // `aria-label`, e não `aria-hidden`: a região tem conteúdo editável.
      role="group"
      aria-label={kind === 'header' ? t('document.band.header') : t('document.band.footer')}
    >
      {band.rows.length === 0 ? null : <BandGrid band={band} {...parts} />}
      <BandCellPieces pieces={band.left} place="left" {...parts} />
      <BandCellPieces pieces={band.center} place="center" {...parts} />
      <BandCellPieces pieces={band.right} place="right" {...parts} />
    </div>
  )
}

interface BandParts {
  pageLabel: string
  totalPages: number
  onEdit?: ((pid: string, text: string) => void) | undefined
}

/** Uma tabela de verdade: o logotipo mora numa célula mesclada. As bordas vêm resolvidas do leitor. */
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
  ({ pageLabel, totalPages, onEdit }: BandParts) =>
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

    const text = pieceText(piece, pageLabel, totalPages)

    const style: React.CSSProperties = {
      fontWeight: piece.bold ? 700 : undefined,
      fontStyle: piece.italic ? 'italic' : undefined,
      color: piece.color,
      fontSize: piece.fontSize,
      fontFamily: piece.fontFamily,
    }

    // O número da página não tem `w:t` onde guardar o que se digitasse nele.
    if (piece.pid === undefined || onEdit === undefined) {
      return (
        <span key={index} style={style}>
          {text}
        </span>
      )
    }

    return (
      <BandText
        key={index}
        pid={piece.pid}
        text={text}
        raw={piece.text ?? ''}
        style={style}
        onEdit={onEdit}
      />
    )
  }

/**
 * O texto sai no `blur`, e não a cada tecla: mudar a configuração de página
 * redesenha as folhas e levaria o cursor embora. A ortografia vale aqui como no
 * corpo; a preferência é lida da loja.
 */
function BandText({
  pid,
  text,
  raw,
  style,
  onEdit,
}: {
  pid: string
  text: string
  /** Com `{n}` e `{total}`, que voltam ao arquivo como campo. */
  raw: string
  style: React.CSSProperties
  onEdit: (pid: string, text: string) => void
}): React.JSX.Element {
  const host = useRef<HTMLSpanElement>(null)
  const spellcheck = usePreferences((state) => state.preferences.spellcheck)

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
      spellCheck={spellcheck}
      // Com o cursor dentro, os campos aparecem como `{n}` e `{total}`: o número desta folha viraria texto fixo.
      onFocus={() => {
        if (host.current !== null && host.current.textContent !== raw) host.current.textContent = raw
      }}
      onBlur={() => {
        const edited = host.current?.textContent ?? ''
        if (host.current !== null) host.current.textContent = text
        onEdit(pid, edited)
      }}
      // Um `w:t` é uma linha só: Enter fecha a edição.
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault()
          host.current?.blur()
        }
      }}
    />
  )
}

/** Cada parágrafo do arquivo é uma linha: num flex, um `<br>` não gera caixa. */
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
