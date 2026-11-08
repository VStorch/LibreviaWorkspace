import { describe, expect, it } from 'vitest'
import { getExtensionField, resolveExtensions, type AnyExtension } from '@tiptap/core'
import { SHORTCUTS, ShortcutOwner, editorKeyOf } from '@shared/shortcuts.js'
import { buildEditorExtensions } from './editor-extensions.js'

function keysOf(extension: AnyExtension): string[] {
  const addKeyboardShortcuts = getExtensionField<() => Record<string, unknown>>(
    extension,
    'addKeyboardShortcuts',
    { name: extension.name, options: extension.options, storage: {}, editor: {}, type: null, parent: null },
  )
  return addKeyboardShortcuts === undefined ? [] : Object.keys(addKeyboardShortcuts())
}

function normalized(key: string): string {
  return key
    .split('-')
    .map((part) => (part.length === 1 ? part.toLowerCase() : part))
    .join('-')
}

/** The editor binds the same command as the menu, so it does not matter which one hears the key. */
const SAME_COMMAND_IN_BOTH = new Set(['Mod-Enter'])

describe('editor keymap', () => {
  it('leaves the menu accelerators to the menu, since a key the editor handles never reaches it', () => {
    const menuKeys = new Map(
      Object.entries(SHORTCUTS)
        .filter(([, shortcut]) => shortcut.owner === ShortcutOwner.Menu)
        .map(([id, shortcut]) => [editorKeyOf(shortcut), id]),
    )

    const collisions = resolveExtensions(buildEditorExtensions(() => {})).flatMap((extension) =>
      keysOf(extension)
        .map(normalized)
        .filter((key) => menuKeys.has(key) && !SAME_COMMAND_IN_BOTH.has(key))
        .map((key) => `${key}: ${extension.name} e ${menuKeys.get(key) ?? ''}`),
    )

    expect(collisions).toEqual([])
  })
})
