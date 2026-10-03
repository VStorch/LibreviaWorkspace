import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  LineSpacingKind,
  MAX_LINE_FACTOR,
  MAX_SPACING_PT,
  MIN_LINE_FACTOR,
  TextAlignment,
  isValidParagraphDraft,
  type ParagraphDraft,
} from '@services/document/paragraph-format.js'
import { useT } from '../../i18n.js'
import { paragraphDraftAt } from '../extensions/paragraph-commands.js'
import { ParagraphIndentFields } from './ParagraphIndentFields.js'
import { ParagraphSpacingFields } from './ParagraphSpacingFields.js'
import { isCustomLineSpacing, lineSpacingChoice, lineSpacingFrom } from './paragraph-draft.js'

/** Abre com o que o parágrafo do cursor já tem: é também um jeito de **ler** a formatação. */
export function ParagraphDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [draft, setDraft] = useState<ParagraphDraft>(() => paragraphDraftAt(editor))

  const valid = isValidParagraphDraft(draft)
  const change = <K extends keyof ParagraphDraft>(key: K, value: ParagraphDraft[K]): void =>
    setDraft({ ...draft, [key]: value })

  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  function apply(): void {
    if (!valid) return
    editor.chain().focus().setParagraphFormat(draft).run()
    onClose()

    // De novo depois de fechar: o painel sai da tela e levaria o foco ao corpo da página.
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label={t('document.paragraph.title')}
      // No elemento, e não globalmente, como o painel de localizar.
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('document.paragraph.alignment')}</span>
          <select
            aria-label={t('document.paragraph.alignment')}
            value={draft.align}
            autoFocus
            onChange={(event) => change('align', event.target.value as TextAlignment)}
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
            onChange={(event) => setDraft({ ...draft, ...lineSpacingFrom(event.target.value) })}
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
              onChange={(event) => change('lineSpacingValue', Number(event.target.value))}
            />
          </label>
        )}
      </div>

      <ParagraphSpacingFields draft={draft} onChange={change} />

      <ParagraphIndentFields draft={draft} onChange={change} />

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.keepNext}
          onChange={(event) => change('keepNext', event.target.checked)}
        />
        <span>{t('document.paragraph.keepWithNext')}</span>
      </label>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.keepLines}
          onChange={(event) => change('keepLines', event.target.checked)}
        />
        <span>{t('document.paragraph.keepLinesTogether')}</span>
      </label>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.widowControl}
          onChange={(event) => change('widowControl', event.target.checked)}
        />
        <span>{t('document.paragraph.widowControl')}</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid ? t('document.paragraph.hintValid') : t('document.paragraph.hintInvalid')}
      </p>

      <div className="popover__actions">
        <button
          type="button"
          className="btn"
          onClick={() => setDraft({ ...DEFAULT_PARAGRAPH_DRAFT, align: draft.align })}
        >
          {t('document.common.restoreDefaults')}
        </button>
        <span className="popover__spacer" />
        {/* `preventDefault` no `mousedown`: o botão não toma o foco do documento. */}
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button
          type="button"
          className="btn btn--primary"
          onMouseDown={keepFocus}
          onClick={apply}
          disabled={!valid}
        >
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
