import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/**
 * O documento travado não muda, venha a mudança de onde vier.
 *
 * O `contenteditable="false"` só segura o teclado e o mouse. Os comandos do
 * Tiptap — que é o que a barra de ferramentas, o menu e os diálogos chamam — não
 * perguntam se o editor é editável: negrito na barra, linha de tabela pelo menu
 * ou "Substituir tudo" no painel de busca mudavam o documento aberto em somente
 * leitura, e o status virava "Não salvo".
 *
 * Aqui é o último portão, e não o único: o menu e o botão direito já param em
 * `useEditorCommands`, que nem abre o diálogo de edição. Este filtro pega o que
 * não passa por lá. Só a transação que **muda o documento** é recusada — a da
 * paginação, a de seleção e a das marcas de formatação passam, porque não editam.
 */
export const ReadOnlyGuard = Extension.create({
  name: 'readOnlyGuard',

  addProseMirrorPlugins() {
    const { editor } = this
    return [
      new Plugin({
        key: new PluginKey('readOnlyGuard'),
        filterTransaction: (transaction) => !transaction.docChanged || editor.isEditable,
      }),
    ]
  },
})
