import { MAX_SPACING_PT, type ParagraphDraft } from '@services/document/paragraph-format.js'
import { useT } from '../../i18n.js'
import type { DraftChange } from './paragraph-draft.js'

interface ParagraphSpacingFieldsProps {
  readonly draft: ParagraphDraft
  readonly onChange: DraftChange
}

/** O espaço antes e depois do parágrafo, em pontos — como o Word o mede. */
export function ParagraphSpacingFields({ draft, onChange }: ParagraphSpacingFieldsProps): React.JSX.Element {
  const t = useT()

  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.paragraph.spacingLegend')}</legend>
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>{t('document.paragraph.spacingBefore')}</span>
          <input
            type="number"
            min={0}
            max={MAX_SPACING_PT}
            step={1}
            value={draft.spaceBefore}
            onChange={(event) => onChange('spaceBefore', Number(event.target.value))}
          />
        </label>
        <label className="popover__field popover__field--narrow">
          <span>{t('document.paragraph.spacingAfter')}</span>
          <input
            type="number"
            min={0}
            max={MAX_SPACING_PT}
            step={1}
            value={draft.spaceAfter}
            onChange={(event) => onChange('spaceAfter', Number(event.target.value))}
          />
        </label>
      </div>
    </fieldset>
  )
}
