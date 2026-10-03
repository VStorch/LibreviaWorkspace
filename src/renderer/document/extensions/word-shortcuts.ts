import { Extension } from '@tiptap/core'
import { SHORTCUTS, type EditorShortcutId, editorKeyOf } from '@shared/shortcuts.js'

/** As teclas estão em `@shared/shortcuts`, com as do menu nativo, que as ocupa primeiro. */
export const WordShortcuts = Extension.create({
  name: 'wordShortcuts',

  /** Acima de 100: `Ctrl+E` é também da marca de código do Tiptap, e a primeira extensão decide. */
  priority: 200,

  addKeyboardShortcuts() {
    const align = (value: string) => () => this.editor.commands.setTextAlign(value)
    const lineHeight = (value: string) => () => this.editor.commands.setBlockLineHeight(value)

    // O tipo é fechado: entrada nova na tabela sem comando aqui não compila.
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
