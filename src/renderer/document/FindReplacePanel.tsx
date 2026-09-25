import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useT } from '../i18n.js'
import type { SearchStatus } from './extensions/search-replace.js'

interface FindReplacePanelProps {
  readonly editor: Editor
  readonly status: SearchStatus
  readonly onClose: () => void
}

export function FindReplacePanel({ editor, status, onClose }: FindReplacePanelProps): React.JSX.Element {
  const t = useT()
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
        placeholder={t('document.findReplace.find')}
        aria-label={t('document.findReplace.find')}
        onChange={(event) => setTerm(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') editor.commands.goToMatch(event.shiftKey ? -1 : 1)
          if (event.key === 'Escape') onClose()
        }}
      />

      <span className={noMatches ? 'findbar__count findbar__count--empty' : 'findbar__count'}>
        {term.length === 0
          ? ''
          : noMatches
            ? t('document.findReplace.noMatches')
            : t('document.findReplace.matchCount', { current: status.current, total: status.total })}
      </span>

      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.goToMatch(-1)}
        disabled={status.total === 0}
        title={t('document.findReplace.previousMatch')}
      >
        ↑
      </button>
      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.goToMatch(1)}
        disabled={status.total === 0}
        title={t('document.findReplace.nextMatch')}
      >
        ↓
      </button>

      <input
        type="text"
        className="findbar__input"
        value={replacement}
        placeholder={t('document.findReplace.replaceWith')}
        aria-label={t('document.findReplace.replaceWith')}
        onChange={(event) => setReplacement(event.target.value)}
      />

      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.replaceCurrentMatch(replacement)}
        disabled={status.total === 0}
      >
        {t('document.findReplace.replace')}
      </button>
      <button
        type="button"
        className="btn"
        onClick={() => editor.commands.replaceAllMatches(replacement)}
        disabled={status.total === 0}
      >
        {t('document.findReplace.replaceAll')}
      </button>

      <label className="findbar__toggle" title={t('document.findReplace.caseSensitive')}>
        <input
          type="checkbox"
          checked={caseSensitive}
          onChange={(event) => setCaseSensitive(event.target.checked)}
        />
        Aa
      </label>

      <button type="button" className="btn" onClick={onClose} aria-label={t('document.findReplace.close')}>
        ✕
      </button>
    </div>
  )
}
