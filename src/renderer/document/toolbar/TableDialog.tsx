import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { MAX_TABLE_COLUMNS, MAX_TABLE_ROWS, isValidTableSize } from '@services/document/table-format.js'

/**
 * Inserir tabela, perguntando de que tamanho.
 *
 * O botão da barra inseria uma 3 × 3 fixa. Era o único jeito de pôr uma tabela no
 * documento, e quem precisava de cinco colunas tinha de inserir e depois
 * acrescentar coluna a coluna — sem comando nenhum na interface para isso.
 *
 * Mesmo desenho dos outros diálogos do documento (`ParagraphDialog`): `popover`
 * com rascunho local, `Escape` fecha, `Enter` aplica e "Inserir" só funciona com o
 * formulário válido.
 */
export function TableDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const [rows, setRows] = useState(3)
  const [columns, setColumns] = useState(3)
  const [headerRow, setHeaderRow] = useState(true)

  const valid = isValidTableSize(rows, columns)
  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  function insert(): void {
    if (!valid) return
    editor.chain().focus().insertTable({ rows, cols: columns, withHeaderRow: headerRow }).run()
    onClose()

    // E de novo depois do fechamento, como no diálogo de parágrafo: o painel sai
    // da tela depois do `focus()` da cadeia e leva o foco do documento com ele.
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label="Inserir tabela"
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') insert()
      }}
    >
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>Linhas</span>
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
          <span>Colunas</span>
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
        {/* O que o arquivo guarda é `w:tblHeader`, e é isso que o rótulo promete:
            a linha reaparece no alto de cada página, e não só fica em negrito. */}
        <span>Linha de cabeçalho, repetida em cada página</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid
          ? 'As colunas nascem com a mesma largura, e a divisória se arrasta depois.'
          : `Entre 1 e ${MAX_TABLE_ROWS} linhas e 1 e ${MAX_TABLE_COLUMNS} colunas.`}
      </p>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          Cancelar
        </button>
        <button
          type="button"
          className="btn btn--primary"
          onMouseDown={keepFocus}
          onClick={insert}
          disabled={!valid}
        >
          Inserir
        </button>
      </div>
    </div>
  )
}
