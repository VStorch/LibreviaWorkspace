import { Extension } from '@tiptap/core'
import { SHORTCUTS, type EditorShortcutId, editorKeyOf } from '@shared/shortcuts.js'

/** The keys are in `@shared/shortcuts`, along with the native menu ones, which take them first. */
export const WordShortcuts = Extension.create({
  name: 'wordShortcuts',

  /** Above 100: `Ctrl+E` also belongs to Tiptap's code mark, and the first extension decides. */
  priority: 200,

  addKeyboardShortcuts() {
    const align = (value: string) => () => this.editor.commands.setTextAlign(value)
    const lineHeight = (value: string) => () => this.editor.commands.setBlockLineHeight(value)

    // The type is closed: a new table entry without a command here does not compile.
    const commands: Record<EditorShortcutId, () => boolean> = {
      alignLeft: align('left'),
      alignCenter: align('center'),
      alignRight: align('right'),
      alignJustify: align('justify'),
      lineHeightSingle: lineHeight(''),
      lineHeightOneAndHalf: lineHeight('1.5'),
      lineHeightDouble: lineHeight('2'),
      superscript: () => this.editor.commands.toggleSuperscript(),
      subscript: () => this.editor.commands.toggleSubscript(),
    }

    const keymap: Record<string, () => boolean> = {}
    for (const id of Object.keys(commands) as EditorShortcutId[]) {
      keymap[editorKeyOf(SHORTCUTS[id])] = commands[id]
    }

    return keymap
  },
})
