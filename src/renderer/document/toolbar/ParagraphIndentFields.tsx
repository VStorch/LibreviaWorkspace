import { FirstLineKind, MAX_INDENT_MM, type ParagraphDraft } from '@services/document/paragraph-format.js'
import { useT } from '../../i18n.js'
import type { DraftChange } from './paragraph-draft.js'

interface ParagraphIndentFieldsProps {
  readonly draft: ParagraphDraft
  readonly onChange: DraftChange
}

/** Os recuos, em milímetros, e o tratamento da primeira linha. */
export function ParagraphIndentFields({ draft, onChange }: ParagraphIndentFieldsProps): React.JSX.Element {
  const t = useT()

  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.paragraph.indentLegend')}</legend>
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>{t('document.paragraph.indentLeft')}</span>
          <input
            type="number"
            min={0}
            max={MAX_INDENT_MM}
            step={1}
            value={draft.indentLeftMm}
            onChange={(event) => onChange('indentLeftMm', Number(event.target.value))}
          />
        </label>
        <label className="popover__field popover__field--narrow">
          <span>{t('document.paragraph.indentRight')}</span>
          <input
            type="number"
            min={0}
            max={MAX_INDENT_MM}
            step={1}
            value={draft.indentRightMm}
            onChange={(event) => onChange('indentRightMm', Number(event.target.value))}
          />
        </label>

        <label className="popover__field">
          <span>{t('document.paragraph.firstLine')}</span>
          <select
            aria-label={t('document.paragraph.firstLine')}
            value={draft.firstLineKind}
            onChange={(event) => onChange('firstLineKind', event.target.value as FirstLineKind)}
          >
            <option value={FirstLineKind.None}>{t('document.paragraph.firstLineNone')}</option>
            <option value={FirstLineKind.Indent}>{t('document.paragraph.firstLineIndent')}</option>
            <option value={FirstLineKind.Hanging}>{t('document.paragraph.firstLineHanging')}</option>
          </select>
        </label>

        {draft.firstLineKind !== FirstLineKind.None && (
          <label className="popover__field popover__field--narrow">
            <span>{t('document.paragraph.firstLineBy')}</span>
            <input
              type="number"
              min={0}
              max={MAX_INDENT_MM}
              step={1}
              value={draft.firstLineMm}
              onChange={(event) => onChange('firstLineMm', Number(event.target.value))}
            />
          </label>
        )}
      </div>
    </fieldset>
  )
}
