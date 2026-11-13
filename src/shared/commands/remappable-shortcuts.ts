import type { MessageKey } from '../i18n/index.js'
import { isRemappable } from '../shortcut-overrides.js'
import type { ShortcutId } from '../shortcuts.js'
import { COMMANDS, type CommandMeta } from './registry.js'

export interface RemappableShortcut {
  readonly id: ShortcutId
  readonly labelKey: MessageKey
}

/** View toggles live outside the registry, as preference switches in main. */
const TOGGLES: readonly RemappableShortcut[] = [
  { id: 'readingMode', labelKey: 'view.reading' },
  { id: 'formattingMarks', labelKey: 'menu.view.formattingMarks' },
  { id: 'navigationPane', labelKey: 'menu.view.navigationPane' },
]

/** In menu order, named as the menu names them. */
export const REMAPPABLE_SHORTCUTS: readonly RemappableShortcut[] = [
  ...Object.values<CommandMeta>(COMMANDS).flatMap((meta) =>
    meta.shortcut !== undefined && meta.registerAccelerator !== false && isRemappable(meta.shortcut)
      ? [{ id: meta.shortcut, labelKey: meta.labelKey }]
      : [],
  ),
  ...TOGGLES,
]
