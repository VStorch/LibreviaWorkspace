import { useState } from 'react'
import { REMAPPABLE_SHORTCUTS } from '@shared/commands/remappable-shortcuts.js'
import {
  ShortcutProblem,
  keyText,
  shortcutKeyOf,
  shortcutProblem,
  type ShortcutTable,
} from '@shared/shortcut-overrides.js'
import { SHORTCUTS, canonicalKeyOf, shortcutHintOf, type ShortcutId } from '@shared/shortcuts.js'
import type { MessageKey, Vars } from '@shared/i18n/index.js'
import { useT } from '../i18n.js'
import { setPreference, useShortcutTable } from '../state/preferences.js'
import { DialogActions } from './DialogActions.js'

type Translate = (key: MessageKey, vars?: Vars) => string

const MODIFIER_CODES = /^(?:Control|Shift|Alt|Meta|OS)(?:Left|Right)?$/u

function overridesOf(table: ShortcutTable): Record<string, string> {
  const overrides: Record<string, string> = {}
  for (const { id } of REMAPPABLE_SHORTCUTS) {
    if (canonicalKeyOf(table[id].key) !== canonicalKeyOf(SHORTCUTS[id].key))
      overrides[id] = keyText(table[id].key)
  }
  return overrides
}

function problemText(t: Translate, problem: ShortcutProblem, hint: string): string {
  if (problem.kind === ShortcutProblem.NeedsModifier) return t('shell.shortcuts.needsModifier')
  if (problem.kind === ShortcutProblem.Reserved) return t('shell.shortcuts.reserved', { key: hint })
  const holder = REMAPPABLE_SHORTCUTS.find((candidate) => candidate.id === problem.by)
  if (holder === undefined) return t('shell.shortcuts.takenByEditor', { key: hint })
  return t('shell.shortcuts.taken', { key: hint, command: t(holder.labelKey) })
}

interface Capture {
  readonly id: ShortcutId
  readonly problem: string | null
}

/** Options → Keyboard shortcuts: the menu's keys, one at a time, checked against every other key. */
export function ShortcutsDialog({ onClose }: { readonly onClose: () => void }): React.JSX.Element {
  const t = useT()
  const [table, setTable] = useState<ShortcutTable>(useShortcutTable())
  const [capture, setCapture] = useState<Capture | null>(null)

  function press(id: ShortcutId, event: React.KeyboardEvent): void {
    event.preventDefault()
    event.stopPropagation()
    if (event.key === 'Escape' && !event.ctrlKey && !event.altKey) return setCapture(null)
    if (MODIFIER_CODES.test(event.code)) return
    const key = shortcutKeyOf(event)
    if (key === null) return
    const problem = shortcutProblem(table, id, key)
    if (problem !== null)
      return setCapture({ id, problem: problemText(t, problem, shortcutHintOf({ ...table[id], key })) })
    setTable({ ...table, [id]: { ...table[id], key } })
    setCapture(null)
  }

  function apply(): void {
    void setPreference({ shortcuts: overridesOf(table) })
    onClose()
  }

  return (
    <div
      className="popover popover--wide shortcuts"
      role="dialog"
      aria-label={t('shell.shortcuts.title')}
      onKeyDown={(event) => {
        event.stopPropagation()
        if (event.key === 'Escape') onClose()
      }}
    >
      <p className="popover__hint">{t('shell.shortcuts.intro')}</p>
      <table className="shortcuts__table">
        <tbody>
          {REMAPPABLE_SHORTCUTS.map(({ id, labelKey }) => (
            <ShortcutRow
              key={id}
              label={t(labelKey)}
              hint={shortcutHintOf(table[id])}
              changed={canonicalKeyOf(table[id].key) !== canonicalKeyOf(SHORTCUTS[id].key)}
              capture={capture?.id === id ? capture : null}
              onCapture={() => setCapture({ id, problem: null })}
              onPress={(event) => press(id, event)}
              onReset={() => setTable({ ...table, [id]: SHORTCUTS[id] })}
            />
          ))}
        </tbody>
      </table>
      <DialogActions confirmLabel={t('document.common.apply')} onConfirm={apply} onCancel={onClose}>
        <button type="button" className="btn" onClick={() => setTable(SHORTCUTS)}>
          {t('shell.shortcuts.resetAll')}
        </button>
      </DialogActions>
    </div>
  )
}

interface ShortcutRowProps {
  readonly label: string
  readonly hint: string
  readonly changed: boolean
  readonly capture: Capture | null
  readonly onCapture: () => void
  readonly onPress: (event: React.KeyboardEvent) => void
  readonly onReset: () => void
}

function ShortcutRow({
  label,
  hint,
  changed,
  capture,
  onCapture,
  onPress,
  onReset,
}: ShortcutRowProps): React.JSX.Element {
  const t = useT()
  return (
    <tr>
      <th scope="row">{label}</th>
      <td>
        {capture === null ? (
          <button
            type="button"
            className="btn shortcuts__key"
            aria-label={`${label}: ${hint}`}
            onClick={onCapture}
          >
            <kbd>{hint}</kbd>
          </button>
        ) : (
          <button type="button" className="btn shortcuts__key" autoFocus onKeyDown={onPress}>
            {t('shell.shortcuts.press')}
          </button>
        )}
        {capture !== null && capture.problem !== null && (
          <span className="shortcuts__problem" role="alert">
            {capture.problem}
          </span>
        )}
      </td>
      <td>
        <button
          type="button"
          className="btn"
          disabled={!changed}
          onClick={onReset}
          aria-label={t('shell.shortcuts.resetOne', { command: label })}
        >
          {t('shell.shortcuts.reset')}
        </button>
      </td>
    </tr>
  )
}
