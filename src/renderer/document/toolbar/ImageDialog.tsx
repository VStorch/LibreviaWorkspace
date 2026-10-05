import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useT } from '../../i18n.js'
import { applyImageProperties, imageAt } from '../extensions/document-image.js'

/** What `w:docPr/@descr` takes without excess. */
const MAX_ALT_LENGTH = 300

/**
 * Alt text goes to `wp:docPr/@descr` and is what the screen reader announces: Word's "Alt Text".
 * Size is set with the image handles.
 */
export function ImageDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const placed = imageAt(editor)
  const [alt, setAlt] = useState(() => {
    const value = placed?.node.attrs['alt']
    return typeof value === 'string' ? value : ''
  })
  const [align, setAlign] = useState(() => currentAlign(editor))

  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  function apply(): void {
    const target = imageAt(editor)
    if (target !== null) {
      applyImageProperties(editor, target, { alt, align: align === '' ? null : align })
    }
    onClose()
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('document.imageDialog.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <label className="popover__field">
        <span>{t('document.imageDialog.altText')}</span>
        <input
          type="text"
          aria-label={t('document.imageDialog.altText')}
          maxLength={MAX_ALT_LENGTH}
          value={alt}
          autoFocus
          onChange={(event) => setAlt(event.target.value)}
        />
      </label>

      <label className="popover__field">
        <span>{t('document.imageDialog.alignment')}</span>
        <select
          aria-label={t('document.imageDialog.alignmentLabel')}
          value={align}
          onChange={(event) => setAlign(event.target.value)}
        >
          <option value="">{t('document.imageDialog.alignSameAsParagraph')}</option>
          <option value="left">{t('document.imageDialog.alignLeft')}</option>
          <option value="center">{t('document.imageDialog.alignCenter')}</option>
          <option value="right">{t('document.imageDialog.alignRight')}</option>
        </select>
      </label>

      <p className="popover__hint">{t('document.imageDialog.hint')}</p>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onMouseDown={keepFocus} onClick={apply}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}

/** The paragraph's, when it lives in one, or its own, when it is a loose block. */
function currentAlign(editor: Editor): string {
  const placed = imageAt(editor)
  if (placed === null) return ''

  if (placed.paragraphPos !== null) {
    const paragraph = editor.state.doc.nodeAt(placed.paragraphPos)
    const align = paragraph?.attrs['textAlign']
    return typeof align === 'string' ? align : ''
  }

  const align = placed.node.attrs['align']
  return typeof align === 'string' ? align : ''
}
