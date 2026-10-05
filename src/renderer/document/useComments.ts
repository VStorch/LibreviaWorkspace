import { useEffect, useMemo, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { DocumentComment } from '@services/document/model.js'
import { resolveComments } from '@services/document/comments.js'
import { useWorkspace } from '../state/workspace.js'
import { commentAnchorsOf } from './extensions/comment.js'

/** Changes only when a comment comes in or goes out. */
function anchorKey(editor: Editor | null): string {
  if (editor === null || editor.isDestroyed) return ''
  return [...commentAnchorsOf(editor.state.doc).keys()].sort().join('\u0000')
}

/** Redraws only when the set of ends changes, not on every letter; see `resolveComments`. */
export function useComments(editor: Editor | null): {
  readonly comments: readonly DocumentComment[]
  readonly outside: ReadonlySet<string>
} {
  const library = useWorkspace((state) => state.comments)
  const outsideList = useWorkspace((state) => state.commentsOutside)
  const [key, setKey] = useState(() => anchorKey(editor))

  useEffect(() => {
    if (editor === null) return
    setKey(anchorKey(editor))
    const update = ({ transaction }: { transaction: { docChanged: boolean } }): void => {
      if (transaction.docChanged) setKey(anchorKey(editor))
    }
    editor.on('transaction', update)
    return () => {
      editor.off('transaction', update)
    }
  }, [editor])

  const outside = useMemo(() => new Set(outsideList), [outsideList])
  const comments = useMemo(
    () => resolveComments(new Set(key === '' ? [] : key.split('\u0000')), library, outside),
    [key, library, outside],
  )
  return { comments, outside }
}
