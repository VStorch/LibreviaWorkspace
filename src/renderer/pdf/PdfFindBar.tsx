import { useEffect, useState } from 'react'
import { useT } from '../i18n.js'
import { sendToViewer, usePdfViewer } from './viewer-state.js'

/** The editor's find bar, without replacing: a PDF opens for reading only. */
export function PdfFindBar({ onClose }: { readonly onClose: () => void }): React.JSX.Element {
  const t = useT()
  const [term, setTerm] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const found = usePdfViewer((viewer) => viewer.found)

  useEffect(() => {
    sendToViewer({ type: 'find', query: term, again: false, backwards: false, caseSensitive })
  }, [term, caseSensitive])

  // Closing the search must clear the highlights, or they stay on the pages.
  useEffect(
    () => () => {
      sendToViewer({ type: 'find-close' })
      usePdfViewer.setState({ found: null })
    },
    [],
  )

  const step = (backwards: boolean): void =>
    sendToViewer({ type: 'find', query: term, again: true, backwards, caseSensitive })
  const total = found?.total ?? 0
  const noMatches = term.length > 0 && found !== null && total === 0

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
          if (event.key === 'Enter') step(event.shiftKey)
          if (event.key === 'Escape') onClose()
        }}
      />

      <span className={noMatches ? 'findbar__count findbar__count--empty' : 'findbar__count'}>
        {term.length === 0 || found === null
          ? ''
          : noMatches
            ? t('document.findReplace.noMatches')
            : t('document.findReplace.matchCount', { current: found.current, total })}
      </span>

      <button
        type="button"
        className="btn"
        onClick={() => step(true)}
        disabled={total === 0}
        title={t('document.findReplace.previousMatch')}
      >
        ↑
      </button>
      <button
        type="button"
        className="btn"
        onClick={() => step(false)}
        disabled={total === 0}
        title={t('document.findReplace.nextMatch')}
      >
        ↓
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
