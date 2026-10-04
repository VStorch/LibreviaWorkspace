import {
  LineSpacingKind,
  MAX_LINE_FACTOR,
  MAX_SPACING_PT,
  MIN_LINE_FACTOR,
  TextAlignment,
  type ParagraphDraft,
} from '@services/document/paragraph-format.js'
import { useT } from '../../i18n.js'
import {
  isCustomLineSpacing,
  lineSpacingChoice,
  lineSpacingFrom,
  type DraftChange,
} from './paragraph-draft.js'

/** O alinhamento e a entrelinha, iguais no diálogo de parágrafo e no de estilo. */
export function ParagraphLineFields({
  draft,
  onChange,
  onLineSpacing,
  autoFocus = false,
}: {
  readonly draft: ParagraphDraft
  readonly onChange: DraftChange
  /** A entrelinha troca o tipo e o valor juntos. */
  readonly onLineSpacing: (spacing: ReturnType<typeof lineSpacingFrom>) => void
  readonly autoFocus?: boolean
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__row">
      <label className="popover__field">
        <span>{t('document.paragraph.alignment')}</span>
        <select
          aria-label={t('document.paragraph.alignment')}
          value={draft.align}
          autoFocus={autoFocus}
          onChange={(event) => onChange('align', event.target.value as TextAlignment)}
        >
          <option value={TextAlignment.Left}>{t('document.paragraph.alignLeft')}</option>
          <option value={TextAlignment.Center}>{t('document.paragraph.alignCenter')}</option>
          <option value={TextAlignment.Right}>{t('document.paragraph.alignRight')}</option>
          <option value={TextAlignment.Justify}>{t('document.paragraph.alignJustify')}</option>
        </select>
      </label>

      <label className="popover__field">
        <span>{t('document.paragraph.lineSpacing')}</span>
        <select
          aria-label={t('document.paragraph.lineSpacing')}
          value={lineSpacingChoice(draft)}
          onChange={(event) => onLineSpacing(lineSpacingFrom(event.target.value))}
        >
          <option value="single">{t('document.paragraph.spacingSingle')}</option>
          <option value="1.15">1,15</option>
          <option value="1.5">1,5</option>
          <option value="2">{t('document.paragraph.spacingDouble')}</option>
          <option value="multiple">{t('document.paragraph.spacingMultiple')}</option>
          <option value="at-least">{t('document.paragraph.spacingAtLeast')}</option>
        </select>
      </label>

      {/* Um campo desabilitado ao lado de "Simples" só faria clicar nele. */}
      {isCustomLineSpacing(draft) && (
        <label className="popover__field popover__field--narrow">
          <span>
            {draft.lineSpacingKind === LineSpacingKind.AtLeast
              ? t('document.paragraph.points')
              : t('document.paragraph.factor')}
          </span>
          <input
            type="number"
            min={draft.lineSpacingKind === LineSpacingKind.AtLeast ? 1 : MIN_LINE_FACTOR}
            max={draft.lineSpacingKind === LineSpacingKind.AtLeast ? MAX_SPACING_PT : MAX_LINE_FACTOR}
            step={draft.lineSpacingKind === LineSpacingKind.AtLeast ? 1 : 0.05}
            value={draft.lineSpacingValue}
            onChange={(event) => onChange('lineSpacingValue', Number(event.target.value))}
          />
        </label>
      )}
    </div>
  )
}
