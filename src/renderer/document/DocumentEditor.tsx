import { useCallback, useEffect, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import { mmToPx, pageDimensionsMm } from '@services/document/model.js'
import type { DocumentNode } from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'
import { DocumentToolbar } from './toolbar/DocumentToolbar.js'
import { FindReplacePanel } from './FindReplacePanel.js'
import { PageSetupPanel } from './PageSetupPanel.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { onEditorCommand } from './editor-commands.js'
import type { SearchStatus } from './extensions/search-replace.js'

/**
 * Editor de documentos.
 *
 * A moldura de página é visual: largura e margens em milímetros, rolagem
 * contínua. O editor não pagina ao vivo — essa decisão está registrada em
 * docs/00-plano-tecnico.md §6.3, e a paginação real acontece na exportação
 * para PDF, na Fase 3.
 */
export function DocumentEditor(): React.JSX.Element {
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const page = useWorkspace((state) => state.page)
  const markDirty = useWorkspace((state) => state.markDirty)
  const setStats = useWorkspace((state) => state.setStats)
  const registerDocumentSource = useWorkspace((state) => state.registerDocumentSource)

  const [searchStatus, setSearchStatus] = useState<SearchStatus>({ total: 0, current: 0 })
  const [findOpen, setFindOpen] = useState(false)
  const [pageSetupOpen, setPageSetupOpen] = useState(false)

  const handleSearchStatus = useCallback((status: SearchStatus) => setSearchStatus(status), [])

  const editor = useEditor({
    extensions: buildEditorExtensions(handleSearchStatus),
    content: initialDoc,
    onUpdate: ({ editor: current }) => {
      markDirty()
      setStats({
        characters: current.storage['characterCount'].characters(),
        words: current.storage['characterCount'].words(),
      })
    },
    onCreate: ({ editor: current }) => {
      setStats({
        characters: current.storage['characterCount'].characters(),
        words: current.storage['characterCount'].words(),
      })
    },
    editorProps: {
      attributes: { class: 'page__content', spellcheck: 'false' },
    },
  })

  // Salvar precisa do conteúdo atual, que só o editor conhece.
  useEffect(() => {
    if (editor === null) return undefined
    registerDocumentSource(() => editor.getJSON() as DocumentNode)
    return () => registerDocumentSource(null)
  }, [editor, registerDocumentSource])

  useEffect(
    () =>
      onEditorCommand((command) => {
        if (command === 'find-replace') setFindOpen(true)
        if (command === 'page-setup') setPageSetupOpen(true)
        if (command === 'insert-page-break') editor?.chain().focus().setPageBreak().run()
      }),
    [editor],
  )

  if (editor === null) return <div className="editor-shell" />

  const { width } = pageDimensionsMm(page)

  return (
    <div className="editor-shell">
      <DocumentToolbar
        editor={editor}
        onOpenFind={() => setFindOpen(true)}
        onOpenPageSetup={() => setPageSetupOpen(true)}
      />

      {findOpen && (
        <FindReplacePanel editor={editor} status={searchStatus} onClose={() => setFindOpen(false)} />
      )}

      {pageSetupOpen && <PageSetupPanel onClose={() => setPageSetupOpen(false)} />}

      <div className="editor-scroll">
        <div
          className="page"
          style={{
            width: `${mmToPx(width)}px`,
            paddingTop: `${mmToPx(page.margins.top)}px`,
            paddingRight: `${mmToPx(page.margins.right)}px`,
            paddingBottom: `${mmToPx(page.margins.bottom)}px`,
            paddingLeft: `${mmToPx(page.margins.left)}px`,
          }}
        >
          <EditorContent editor={editor} />
        </div>
      </div>
    </div>
  )
}
