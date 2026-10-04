import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { MessageKey } from '@shared/i18n/index.js'
import { contentWidthMm } from '@services/document/model.js'
import { mmToPx } from '@services/units.js'
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
import { DialogActions } from '../../components/DialogActions.js'
import { useT } from '../../i18n.js'
import { applyTableDraft, tablePlacementAt } from '../extensions/table-look.js'
import { useWorkspace } from '../../state/workspace.js'

/** Só o que o gravador leva ao `.docx`. */
const BORDER_STYLE_KEYS: readonly { readonly value: CellBorderStyle; readonly labelKey: MessageKey }[] = [
  { value: CellBorderStyle.Single, labelKey: 'document.tableProperties.borderSingle' },
  { value: CellBorderStyle.Double, labelKey: 'document.tableProperties.borderDouble' },
  { value: CellBorderStyle.Dashed, labelKey: 'document.tableProperties.borderDashed' },
  { value: CellBorderStyle.Dotted, labelKey: 'document.tableProperties.borderDotted' },
  { value: CellBorderStyle.None, labelKey: 'document.tableProperties.borderNone' },
]

const SIDE_LABEL_KEYS: Record<CellBorderSide, MessageKey> = {
  top: 'document.tableProperties.sideTop',
  right: 'document.tableProperties.sideRight',
  bottom: 'document.tableProperties.sideBottom',
  left: 'document.tableProperties.sideLeft',
}

/**
 * **Só o que o arquivo guarda**: `w:tblGrid`, `w:tcBorders`, `w:shd` e
 * `w:tblHeader`. Margem interna, direção do texto e alinhamento vertical o gravador
 * não leva, e oferecê-los prometeria o que se perde ao salvar.
 */
export function TablePropertiesDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const page = useWorkspace((state) => state.page)
  const contentWidthPx = mmToPx(contentWidthMm(page))

  const [draft, setDraft] = useState<TableDraft>(
    () => tablePlacementAt(editor, contentWidthPx)?.draft ?? DEFAULT_TABLE_DRAFT,
  )

  const valid = isValidTableDraft(draft)

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
      aria-label={t('table.properties')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <BorderFields draft={draft} onChange={change} />

      <ShadingFields draft={draft} onChange={change} />

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.headerRow}
          onChange={(event) => change('headerRow', event.target.checked)}
        />
        <span>{t('document.tableProperties.repeatHeader')}</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid ? t('document.tableProperties.hintValid') : t('document.tableProperties.hintInvalid')}
      </p>

      <DialogActions
        confirmLabel={t('document.common.apply')}
        onConfirm={apply}
        onCancel={onClose}
        disabled={!valid}
        keepEditorFocus
      >
        <button
          type="button"
          className="btn"
          onClick={() => setDraft({ ...DEFAULT_TABLE_DRAFT, columnWidthMm: draft.columnWidthMm })}
        >
          {t('document.common.restoreDefaults')}
        </button>
      </DialogActions>
    </div>
  )
}

type TableDraftChange = <K extends keyof TableDraft>(key: K, value: TableDraft[K]) => void

/** A largura das colunas e a borda: estilo, espessura, cor e lados. */
function BorderFields({
  draft,
  onChange,
}: {
  draft: TableDraft
  onChange: TableDraftChange
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>{t('document.tableProperties.columnWidth')}</span>
          <input
            type="number"
            min={MIN_COLUMN_WIDTH_MM}
            max={MAX_COLUMN_WIDTH_MM}
            step={1}
            value={draft.columnWidthMm ?? ''}
            autoFocus
            onChange={(event) => onChange('columnWidthMm', Number(event.target.value))}
          />
        </label>

        <label className="popover__field">
          <span>{t('document.tableProperties.border')}</span>
          <select
            aria-label={t('document.tableProperties.borderStyle')}
            value={draft.borderStyle}
            onChange={(event) => onChange('borderStyle', event.target.value as CellBorderStyle)}
          >
            {BORDER_STYLE_KEYS.map((style) => (
              <option key={style.value} value={style.value}>
                {t(style.labelKey)}
              </option>
            ))}
          </select>
        </label>

        <label className="popover__field popover__field--narrow">
          <span>{t('document.tableProperties.borderWidth')}</span>
          <input
            type="number"
            min={MIN_BORDER_PT}
            max={MAX_BORDER_PT}
            step={0.25}
            value={draft.borderWidthPt}
            onChange={(event) => onChange('borderWidthPt', Number(event.target.value))}
          />
        </label>

        <label className="popover__field popover__field--narrow">
          <span>{t('document.tableProperties.borderColor')}</span>
          <input
            type="color"
            value={draft.borderColor}
            onChange={(event) => onChange('borderColor', event.target.value)}
          />
        </label>
      </div>

      <div className="popover__row">
        {CELL_BORDER_SIDES.map((side) => (
          <label className="popover__check" key={side}>
            <input
              type="checkbox"
              checked={draft.sides[side]}
              onChange={(event) => onChange('sides', { ...draft.sides, [side]: event.target.checked })}
            />
            <span>{t(SIDE_LABEL_KEYS[side])}</span>
          </label>
        ))}
      </div>
    </>
  )
}

function ShadingFields({
  draft,
  onChange,
}: {
  draft: TableDraft
  onChange: TableDraftChange
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__row">
      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.shaded}
          onChange={(event) => onChange('shaded', event.target.checked)}
        />
        <span>{t('document.tableProperties.shading')}</span>
      </label>

      <label className="popover__field popover__field--narrow">
        <span>{t('document.tableProperties.shadingColor')}</span>
        <input
          type="color"
          value={draft.shadingColor}
          disabled={!draft.shaded}
          onChange={(event) => onChange('shadingColor', event.target.value)}
        />
      </label>
    </div>
  )
}
