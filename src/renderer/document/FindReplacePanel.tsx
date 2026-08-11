import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { SearchStatus } from './extensions/search-replace.js'

interface FindReplacePanelProps {
  readonly editor: Editor
  readonly status: SearchStatus
  readonly onClose: () => void
}

export function FindReplacePanel({ editor, status, onClose }: FindReplacePanelProps): React.JSX.Element {
  const [term, setTerm] = useState('')
  const [replacement, setReplacement] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)

  useEffect(() => {
    editor.commands.setSearchTerm(term, caseSensitive)
  }, [editor, term, caseSensitive])

  // Fechar a busca precisa apagar os destaques, senão eles ficam na tela.
  useEffect(() => () => void editor.commands.clearSearch(), [editor])

  const noMatches = term.length > 0 && status.total === 0

  return (
    <div className="findbar" role="search">
      <input
        type="text"
        className="findbar__input"
        value={term}
        autoFocus
        placeholder="Localizar"
        aria-label="Localizar"
        onChange={(event) => setTerm(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') editor.commands.goToMatch(event.shiftKey ? -1 : 1)
          if (event.key === 'Escape') onClose()
        }}
      />

      <span className={noMatches ? 'findbar__count findbar__count--empty' : 'findbar__count'}>
        {term.length === 0 ? '' : noMatches ? 'nenhuma' : `${status.current} de ${status.total}`}
      </span>

      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.goToMatch(-1)}
        disabled={status.total === 0}
        title="Ocorrência anterior (Shift+Enter)"
      >
        ↑
      </button>
      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.goToMatch(1)}
        disabled={status.total === 0}
        title="Próxima ocorrência (Enter)"
      >
        ↓
      </button>

      <input
        type="text"
        className="findbar__input"
        value={replacement}
        placeholder="Substituir por"
        aria-label="Substituir por"
        onChange={(event) => setReplacement(event.target.value)}
      />

      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.replaceCurrentMatch(replacement)}
        disabled={status.total === 0}
      >
        Substituir
      </button>
      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.replaceAllMatches(replacement)}
        disabled={status.total === 0}
      >
        Tudo
      </button>

      <label className="findbar__toggle" title="Diferenciar maiúsculas de minúsculas">
        <input
          type="checkbox"
          checked={caseSensitive}
          onChange={(event) => setCaseSensitive(event.target.checked)}
        />
        Aa
      </label>

      <button type="button" className="btn" onClick={onClose} aria-label="Fechar busca">
        ✕
      </button>
    </div>
  )
}
