import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { DEFAULT_COLUMN_SPACING_MM, allSections, withColumns } from '@services/document/sections.js'
import { DialogActions } from '../components/DialogActions.js'
import { useT } from '../i18n.js'
import { commitSections, resolvedOf, sectionAtCursor } from './section-commands.js'
import { SectionScopeChoice, type SectionScope } from './SectionScopeChoice.js'

const MAX_COLUMNS = 12
const MAX_SPACING_MM = 100

/** For the cursor's section or the whole document, like page setup. */
export function ColumnsDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [{ page, sections, bodyId }] = useState(() => resolvedOf(editor.state.doc))
  const [index] = useState(() => sectionAtCursor(editor))
  const current = allSections(page, sections)[index] ?? page
  const [count, setCount] = useState(current.columns?.count ?? 1)
  const [spaceMm, setSpaceMm] = useState(current.columns?.spaceMm ?? DEFAULT_COLUMN_SPACING_MM)
  const [separator, setSeparator] = useState(current.columns?.separator ?? false)
  const [scope, setScope] = useState<SectionScope>('section')

  const valid =
    Number.isInteger(count) && count >= 1 && count <= MAX_COLUMNS && spaceMm >= 0 && spaceMm <= MAX_SPACING_MM

  function apply(): void {
    if (!valid) return
    commitSections(withColumns({ page, sections }, index, { count, spaceMm, separator }, scope), bodyId)
    onClose()
    requestAnimationFrame(() => editor.view.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label={t('document.columns.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field popover__field--narrow">
          <span>{t('document.columns.count')}</span>
          <input
            type="number"
            min={1}
            max={MAX_COLUMNS}
            step={1}
            value={count}
            autoFocus
            onChange={(event) => setCount(Math.round(Number(event.target.value)))}
          />
        </label>
        <label className="popover__field popover__field--narrow">
          <span>{t('document.columns.spacing')}</span>
          <input
            type="number"
            min={0}
            max={MAX_SPACING_MM}
            step={0.5}
            value={spaceMm}
            disabled={count <= 1}
            onChange={(event) => setSpaceMm(Number(event.target.value))}
          />
        </label>
      </div>
      <label className="popover__check">
        <input
          type="checkbox"
          checked={separator}
          disabled={count <= 1}
          onChange={(event) => setSeparator(event.target.checked)}
        />
        {t('document.columns.separator')}
      </label>
      {sections.length > 0 && <SectionScopeChoice name="columns-scope" scope={scope} onChange={setScope} />}
      <DialogActions
        confirmLabel={t('document.common.apply')}
        onConfirm={apply}
        onCancel={onClose}
        disabled={!valid}
      />
    </div>
  )
}
