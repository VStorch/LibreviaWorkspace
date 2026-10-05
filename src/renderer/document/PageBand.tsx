import { useEffect, useRef } from 'react'
import { linesOf, pieceText, type Band, type BandCell, type BandPiece } from '@services/document/band.js'
import { usePreferences } from '../state/preferences.js'
import { useT } from '../i18n.js'

/**
 * **The text is editable; the frame is not.** Each piece with a file address takes the cursor, and
 * what is typed goes back to its `w:t`; the rest is the drawing of an OOXML part that goes back
 * intact. `pointer-events: none` on the band, so a click in the margin does not pull the cursor out
 * of the body.
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
  /** In the `w:pgNumType` format. */
  pageLabel: string
  totalPages: number
  /** Half the margin: a corporate header is wider than the text column. */
  insetPx: number
  /**
   * `w:pgMar/@header` and `@footer`: drawing in one place and counting from another would push the
   * body down wrongly.
   */
  offsetPx: number
  /** Absent when the document is locked: no piece takes the cursor. */
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
      // `aria-label`, not `aria-hidden`: the region has editable content.
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

/** A real table: the logo lives in a merged cell. Borders come resolved from the reader. */
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

    // The page number has no `w:t` to store typed text.
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
 * Text goes out on `blur`, not on every key press: changing the page setup redraws the sheets and
 * would take the cursor away. Spelling applies here as in the body; the preference is read from the
 * store.
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
  /** With `{n}` and `{total}`, which go back to the file as fields. */
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
      // With the cursor inside, fields show as `{n}` and `{total}`: this sheet's number would
      // become fixed text.
      onFocus={() => {
        if (host.current !== null && host.current.textContent !== raw) host.current.textContent = raw
      }}
      onBlur={() => {
        const edited = host.current?.textContent ?? ''
        if (host.current !== null) host.current.textContent = text
        onEdit(pid, edited)
      }}
      // A `w:t` is a single line: Enter ends editing.
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault()
          host.current?.blur()
        }
      }}
    />
  )
}

/** Each file paragraph is a line: in a flex container, a `<br>` creates no box. */
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
