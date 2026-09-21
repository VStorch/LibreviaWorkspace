import { Extension } from '@tiptap/core'

/**
 * Os atalhos do Word que o editor não tinha.
 *
 * Num lugar só, e não espalhados por cada extensão, porque o que decide cada
 * tecla aqui não é a extensão: é **o que já está ocupado**. A lista abaixo é o
 * resultado dessa conferência, e mantê-la junta é o que impede a próxima
 * adição de roubar uma tecla em silêncio.
 *
 * ## O que já existia, e por isso não está aqui
 *
 * | Tecla                | Quem a tem            | O que faz            |
 * | -------------------- | --------------------- | -------------------- |
 * | `Ctrl+B/I/U`         | StarterKit            | negrito, itálico, sublinhado |
 * | `Ctrl+Alt+1`…`6`     | `Heading`             | títulos — é o atalho do Word |
 * | `Ctrl+Alt+0`         | `Paragraph`           | corpo de texto       |
 * | `Ctrl+Shift+L/E/R/J` | `TextAlign`           | alinhamento (padrão do Tiptap) |
 * | `Ctrl+]` / `Ctrl+[`  | `Indent`              | recuo                |
 * | `Ctrl+.` / `Ctrl+,`  | `Superscript`/`Subscript` | sobrescrito e subscrito |
 *
 * ## Duas colisões, e como foram resolvidas
 *
 * **`Ctrl+E` é do Word e do `Code`.** No Word centraliza; no Tiptap liga a marca
 * de código embutido, que esta barra de ferramentas nem oferece. Centralizar
 * ganha, e ganha pela `priority`: o Tiptap monta os plugins de teclado na ordem
 * das extensões, e a primeira que trata a tecla decide. Prioridade acima da
 * padrão (100) põe este mapa na frente. A marca de código continua alcançável
 * pela regra de entrada de crase.
 *
 * **`Ctrl+Shift+N` é do Word e do menu.** No Word volta o parágrafo para corpo
 * de texto; aqui é "nova planilha", e acelerador de menu é registrado no
 * processo main — ele intercepta a tecla antes de o renderer vê-la, e um atalho
 * daqui para ela simplesmente nunca rodaria. O menu fica com a tecla, que já
 * está documentada, e o corpo de texto segue no `Ctrl+Alt+0` que a extensão
 * `Paragraph` já dá — vizinho natural do `Ctrl+Alt+1`…`6` dos títulos.
 */

export const WordShortcuts = Extension.create({
  name: 'wordShortcuts',

  // Acima do padrão de propósito: ver a nota sobre `Ctrl+E` no comentário do
  // arquivo. Sem isto a marca de código venceria e o texto não centralizaria.
  priority: 200,

  addKeyboardShortcuts() {
    const align = (value: string) => () => this.editor.commands.setTextAlign(value)
    const lineHeight = (value: string) => () => this.editor.commands.setBlockLineHeight(value)

    return {
      // Alinhamento, como no Word e no Writer.
      'Mod-l': align('left'),
      'Mod-e': align('center'),
      'Mod-r': align('right'),
      'Mod-j': align('justify'),

      /**
       * Entrelinha: `Ctrl+1` simples, `Ctrl+5` um e meio, `Ctrl+2` duplo.
       *
       * São os do Word, e é por isso que os títulos ficam no `Ctrl+Alt+1`…`6`:
       * no Word `Ctrl+1` nunca foi "Título 1". A medida é dita em **linhas**, como
       * no Word: vazio é o espaçamento simples, e quem traduz para a medida do CSS
       * — que depende da altura natural da fonte — é `paragraph-format`.
       */
      'Mod-1': lineHeight(''),
      'Mod-5': lineHeight('1.5'),
      'Mod-2': lineHeight('2'),

      /**
       * Sobrescrito e subscrito, como no Word.
       *
       * `Ctrl+Shift+=` e `Ctrl+=`. O `=` sai pelo código da tecla, e não pelo
       * caractere: com Shift o navegador informa `+`, e é o `prosemirror-keymap`
       * que desfaz isso ao tentar o nome derivado do `keyCode`.
       *
       * Os padrões do Tiptap — `Ctrl+.` e `Ctrl+,` — continuam valendo, e não
       * por acaso: são a saída para o teclado em que o `=` não é uma tecla só.
       */
      'Mod-Shift-=': () => this.editor.commands.toggleSuperscript(),
      'Mod-=': () => this.editor.commands.toggleSubscript(),
    }
  },
})
