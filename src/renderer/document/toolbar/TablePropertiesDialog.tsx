import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { contentWidthMm, mmToPx } from '@services/document/model.js'
import {
  CELL_BORDER_SIDES,
  CellBorderStyle,
  DEFAULT_TABLE_DRAFT,
  MAX_BORDER_PT,
  MAX_COLUMN_WIDTH_MM,
  MIN_BORDER_PT,
  MIN_COLUMN_WIDTH_MM,
  isValidTableDraft,
  type CellBorderSide,
  type TableDraft,
} from '@services/document/table-format.js'
import { applyTableDraft, tablePlacementAt } from '../extensions/table-look.js'
import { useWorkspace } from '../../state/workspace.js'

/** Só o que o gravador leva ao `.docx` — o resto não é oferecido. */
const BORDER_STYLES: readonly { readonly value: CellBorderStyle; readonly label: string }[] = [
  { value: CellBorderStyle.Single, label: 'Simples' },
  { value: CellBorderStyle.Double, label: 'Dupla' },
  { value: CellBorderStyle.Dashed, label: 'Tracejada' },
  { value: CellBorderStyle.Dotted, label: 'Pontilhada' },
  { value: CellBorderStyle.None, label: 'Apagada' },
]

const SIDE_LABELS: Record<CellBorderSide, string> = {
  top: 'Acima',
  right: 'À direita',
  bottom: 'Abaixo',
  left: 'À esquerda',
}

/**
 * Propriedades da tabela: largura da coluna, bordas e sombreamento da célula, e a
 * linha de cabeçalho que se repete em cada página.
 *
 * **Só o que o arquivo sabe guardar.** Cada campo aqui tem um destino em OOXML —
 * `w:tblGrid`, `w:tcBorders`, `w:shd` e `w:trPr/w:tblHeader` —, e é por isso que
 * não há margem interna, direção do texto nem alinhamento vertical: o gravador não
 * os leva, e oferecê-los seria prometer o que se perde ao salvar.
 */
export function TablePropertiesDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const page = useWorkspace((state) => state.page)
  const contentWidthPx = mmToPx(contentWidthMm(page))

  const [draft, setDraft] = useState<TableDraft>(
    () => tablePlacementAt(editor, contentWidthPx)?.draft ?? DEFAULT_TABLE_DRAFT,
  )

  const valid = isValidTableDraft(draft)
  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  const change = <K extends keyof TableDraft>(key: K, value: TableDraft[K]): void =>
    setDraft({ ...draft, [key]: value })

  function apply(): void {
    if (!valid) return
    applyTableDraft(editor, draft, contentWidthPx)
    onClose()
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label="Propriedades da tabela"
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>Largura da coluna (mm)</span>
          <input
            type="number"
            min={MIN_COLUMN_WIDTH_MM}
            max={MAX_COLUMN_WIDTH_MM}
            step={1}
            value={draft.columnWidthMm ?? ''}
            autoFocus
            onChange={(event) => change('columnWidthMm', Number(event.target.value))}
          />
        </label>

        <label className="popover__field">
          <span>Borda</span>
          <select
            aria-label="Estilo da borda"
            value={draft.borderStyle}
            onChange={(event) => change('borderStyle', event.target.value as CellBorderStyle)}
          >
            {BORDER_STYLES.map((style) => (
              <option key={style.value} value={style.value}>
                {style.label}
              </option>
            ))}
          </select>
        </label>

        <label className="popover__field popover__field--narrow">
          <span>Espessura (pt)</span>
          <input
            type="number"
            min={MIN_BORDER_PT}
            max={MAX_BORDER_PT}
            step={0.25}
            value={draft.borderWidthPt}
            onChange={(event) => change('borderWidthPt', Number(event.target.value))}
          />
        </label>

        <label className="popover__field popover__field--narrow">
          <span>Cor da borda</span>
          <input
            type="color"
            value={draft.borderColor}
            onChange={(event) => change('borderColor', event.target.value)}
          />
        </label>
      </div>

      <div className="popover__row">
        {CELL_BORDER_SIDES.map((side) => (
          <label className="popover__check" key={side}>
            <input
              type="checkbox"
              checked={draft.sides[side]}
              onChange={(event) => change('sides', { ...draft.sides, [side]: event.target.checked })}
            />
            <span>{SIDE_LABELS[side]}</span>
          </label>
        ))}
      </div>

      <div className="popover__row">
        <label className="popover__check">
          <input
            type="checkbox"
            checked={draft.shaded}
            onChange={(event) => change('shaded', event.target.checked)}
          />
          <span>Sombreamento</span>
        </label>

        <label className="popover__field popover__field--narrow">
          <span>Cor do fundo</span>
          <input
            type="color"
            value={draft.shadingColor}
            disabled={!draft.shaded}
            onChange={(event) => change('shadingColor', event.target.value)}
          />
        </label>
      </div>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.headerRow}
          onChange={(event) => change('headerRow', event.target.checked)}
        />
        <span>Repetir a primeira linha no alto de cada página</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid
          ? 'Borda e sombreamento valem para as células que a seleção tocar; a largura, para a coluna do cursor.'
          : 'Há medida fora da faixa que o documento aceita.'}
      </p>

      <div className="popover__actions">
        <button
          type="button"
          className="btn"
          onClick={() => setDraft({ ...DEFAULT_TABLE_DRAFT, columnWidthMm: draft.columnWidthMm })}
        >
          Restaurar padrão
        </button>
        <span className="popover__spacer" />
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          Cancelar
        </button>
        <button
          type="button"
          className="btn btn--primary"
          onMouseDown={keepFocus}
          onClick={apply}
          disabled={!valid}
        >
          Aplicar
        </button>
      </div>
    </div>
  )
}
