import { Extension } from '@tiptap/core'
import { SHORTCUTS, type EditorShortcutId, editorKeyOf } from '@shared/shortcuts.js'

/**
 * Os atalhos do Word que o editor não tinha.
 *
 * As teclas não estão aqui: estão em `@shared/shortcuts`, junto com as do menu
 * nativo, porque quem decide cada tecla não é a extensão — é **o que já está
 * ocupado**, e o menu ocupa antes de o renderer ver o evento. Este arquivo diz
 * apenas o que cada atalho faz; a combinação vem da tabela, e o teste dela é o que
 * impede a próxima adição de roubar uma tecla em silêncio.
 */
export const WordShortcuts = Extension.create({
  name: 'wordShortcuts',

  /**
   * Acima do padrão (100) de propósito.
   *
   * `Ctrl+E` é do Word e da marca de código do Tiptap: o Tiptap monta os plugins de
   * teclado na ordem das extensões, e a primeira que trata a tecla decide. Sem esta
   * prioridade a marca de código venceria e o texto não centralizaria.
   */
  priority: 200,

  addKeyboardShortcuts() {
    const align = (value: string) => () => this.editor.commands.setTextAlign(value)
    const lineHeight = (value: string) => () => this.editor.commands.setBlockLineHeight(value)

    // Um comando para cada atalho que a tabela dá como nosso. O tipo do registro é
    // fechado: entrada nova na tabela sem comando aqui não compila.
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
