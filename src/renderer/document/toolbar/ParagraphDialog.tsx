import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  isValidParagraphDraft,
  type ParagraphDraft,
} from '@services/document/paragraph-format.js'
import { DialogActions } from '../../components/DialogActions.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { useT } from '../../i18n.js'
import { paragraphDraftAt } from '../extensions/paragraph-commands.js'
import { ParagraphIndentFields } from './ParagraphIndentFields.js'
import { ParagraphSpacingFields } from './ParagraphSpacingFields.js'
import { ParagraphLineFields } from './ParagraphLineFields.js'

const KEEP_OPTIONS = [
  ['keepNext', 'document.paragraph.keepWithNext'],
  ['keepLines', 'document.paragraph.keepLinesTogether'],
  ['widowControl', 'document.paragraph.widowControl'],
] as const satisfies readonly (readonly [keyof ParagraphDraft, MessageKey])[]

/**
 * Opens with what the cursor's paragraph already has: it is also a way to **read** the formatting.
 */
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

  function apply(): void {
    if (!valid) return
    editor.chain().focus().setParagraphFormat(draft).run()
    onClose()

    // Again after closing: the panel leaves the screen and would take focus to the page body.
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label={t('document.paragraph.title')}
      // On the element, not globally, like the find panel.
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <ParagraphLineFields
        draft={draft}
        onChange={change}
        onLineSpacing={(spacing) => setDraft({ ...draft, ...spacing })}
        autoFocus
      />
      <ParagraphSpacingFields draft={draft} onChange={change} />
      <ParagraphIndentFields draft={draft} onChange={change} />
      {KEEP_OPTIONS.map(([key, label]) => (
        <label key={key} className="popover__check">
          <input
            type="checkbox"
            checked={draft[key]}
            onChange={(event) => change(key, event.target.checked)}
          />
          <span>{t(label)}</span>
        </label>
      ))}
      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid ? t('document.paragraph.hintValid') : t('document.paragraph.hintInvalid')}
      </p>
      {/* `preventDefault` on `mousedown`: the button does not take focus from the document. */}
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
          onClick={() => setDraft({ ...DEFAULT_PARAGRAPH_DRAFT, align: draft.align })}
        >
          {t('document.common.restoreDefaults')}
        </button>
      </DialogActions>
    </div>
  )
}
