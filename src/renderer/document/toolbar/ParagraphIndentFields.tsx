import { FirstLineKind, MAX_INDENT_MM, type ParagraphDraft } from '@services/document/paragraph-format.js'
import type { DraftChange } from './paragraph-draft.js'

interface ParagraphIndentFieldsProps {
  readonly draft: ParagraphDraft
  readonly onChange: DraftChange
}

/** Os recuos, em milímetros, e o tratamento da primeira linha. */
export function ParagraphIndentFields({ draft, onChange }: ParagraphIndentFieldsProps): React.JSX.Element {
  return (
    <fieldset className="popover__fieldset">
      <legend>Recuo (mm)</legend>
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>Esquerda</span>
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
          <span>Direita</span>
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
          <span>Primeira linha</span>
          <select
            aria-label="Primeira linha"
            value={draft.firstLineKind}
            onChange={(event) => onChange('firstLineKind', event.target.value as FirstLineKind)}
          >
            <option value={FirstLineKind.None}>Nenhum</option>
            <option value={FirstLineKind.Indent}>Recuo</option>
            <option value={FirstLineKind.Hanging}>Deslocamento</option>
          </select>
        </label>

        {draft.firstLineKind !== FirstLineKind.None && (
          <label className="popover__field popover__field--narrow">
            <span>Em</span>
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
