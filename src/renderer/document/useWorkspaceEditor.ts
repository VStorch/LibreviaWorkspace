import { useCallback, useState } from 'react'
import { useEditor, type Editor } from '@tiptap/react'
import { currentPreferences } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import { t as translateNow } from '../i18n.js'
import { buildEditorExtensions } from './editor-extensions.js'
import { isPaginationOnly } from './extensions/pagination.js'
import type { SearchStatus } from './extensions/search-replace.js'

export interface WorkspaceEditor {
  readonly editor: Editor | null
  readonly searchStatus: SearchStatus
  /** Sobe a cada edição de verdade: é o que faz a paginação medir de novo. */
  readonly contentRevision: number
  readonly touch: () => void
}

/** Da loja, e não das props: o editor é criado uma vez, e o resto chega pelos efeitos. */
export function useWorkspaceEditor(): WorkspaceEditor {
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const markDirty = useWorkspace((state) => state.markDirty)
  const setStats = useWorkspace((state) => state.setStats)
  const [contentRevision, setContentRevision] = useState(0)
  const [searchStatus, setSearchStatus] = useState<SearchStatus>({ total: 0, current: 0 })
  const handleSearchStatus = useCallback((status: SearchStatus) => setSearchStatus(status), [])
  const touch = useCallback(() => setContentRevision((value) => value + 1), [])

  const editor = useEditor({
    extensions: buildEditorExtensions(handleSearchStatus, {
      isTypographyEnabled: () => currentPreferences().typography,
      invisibleCharactersVisible: currentPreferences().invisibleCharacters,
      isKnownComment: (cid) => useWorkspace.getState().comments.some((comment) => comment.id === cid),
      notes: () => useWorkspace.getState().notes,
      isTrackingChanges: () => useWorkspace.getState().trackChanges === true,
      // O `w:author` é obrigatório: sem nome, um autor genérico, como no Word.
      revisionAuthor: () => currentPreferences().authorName.trim() || translateNow('revisions.unknownAuthor'),
    }),
    content: initialDoc,
    onUpdate: ({ editor: current, transaction }) => {
      // A paginação chega como transação: tratá-la como edição sujaria o documento.
      if (isPaginationOnly(transaction)) return

      markDirty()
      touch()
      setStats(statsOf(current))
    },
    onCreate: ({ editor: current }) => setStats(statsOf(current)),
    editorProps: {
      attributes: {
        class: 'page__content',
        spellcheck: currentPreferences().spellcheck ? 'true' : 'false',
      },
    },
  })

  return { editor, searchStatus, contentRevision, touch }
}

function statsOf(editor: Editor): { characters: number; words: number } {
  return {
    characters: editor.storage['characterCount'].characters(),
    words: editor.storage['characterCount'].words(),
  }
}
