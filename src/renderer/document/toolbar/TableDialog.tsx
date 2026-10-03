import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { MAX_TABLE_COLUMNS, MAX_TABLE_ROWS, isValidTableSize } from '@services/document/table-format.js'
import { useT } from '../../i18n.js'

export function TableDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [rows, setRows] = useState(3)
  const [columns, setColumns] = useState(3)
  const [headerRow, setHeaderRow] = useState(true)

  const valid = isValidTableSize(rows, columns)
  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  function insert(): void {
    if (!valid) return
    editor.chain().focus().insertTable({ rows, cols: columns, withHeaderRow: headerRow }).run()
    onClose()

    // De novo depois de fechar, como no diálogo de parágrafo.
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('document.insert.table')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') insert()
      }}
    >
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>{t('document.tableDialog.rows')}</span>
          <input
            type="number"
            min={1}
            max={MAX_TABLE_ROWS}
            step={1}
            value={rows}
            autoFocus
            onChange={(event) => setRows(Math.trunc(Number(event.target.value)))}
          />
        </label>

        <label className="popover__field popover__field--narrow">
          <span>{t('document.tableDialog.columns')}</span>
          <input
            type="number"
            min={1}
            max={MAX_TABLE_COLUMNS}
            step={1}
            value={columns}
            onChange={(event) => setColumns(Math.trunc(Number(event.target.value)))}
          />
        </label>
      </div>

      <label className="popover__check">
        <input type="checkbox" checked={headerRow} onChange={(event) => setHeaderRow(event.target.checked)} />
        {/* `w:tblHeader`: a linha reaparece no alto de cada página. */}
        <span>{t('document.tableDialog.headerRow')}</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid
          ? t('document.tableDialog.hintValid')
          : t('document.tableDialog.hintInvalid', { maxRows: MAX_TABLE_ROWS, maxCols: MAX_TABLE_COLUMNS })}
      </p>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button
          type="button"
          className="btn btn--primary"
          onMouseDown={keepFocus}
          onClick={insert}
          disabled={!valid}
        >
          {t('document.tableDialog.insert')}
        </button>
      </div>
    </div>
  )
}
