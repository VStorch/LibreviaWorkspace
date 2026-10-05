import { useEffect, useLayoutEffect, useMemo, useRef, useState, type RefObject } from 'react'
import type { Editor } from '@tiptap/react'
import { latexToMathMl } from '@services/document/latex.js'
import { mathMlToLatex } from '@services/document/mathml-latex.js'
import { sanitizeMathMl, type MathElement } from '@services/document/mathml.js'
import { MATH_PALETTE, type MathTemplate } from '@services/document/math-palette.js'
import { buildMath } from './extensions/math.js'
import { insertEquation, replaceEquation, type EquationTarget } from './math-commands.js'
import { useT } from '../i18n.js'

/**
 * LaTeX is the source; Temml draws it (`latex.ts`), through the same filter as file equations. A
 * `.docx` equation gets LaTeX on open (`mathml-latex.ts`), and OK without changes does not touch
 * the document. A locked document opens it for viewing only.
 */
export function MathDialog({
  editor,
  target,
  readOnly,
  onClose,
}: {
  readonly editor: Editor
  readonly target: EquationTarget
  readonly readOnly: boolean
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()

  // Read once, on open.
  const [initial] = useState(() => initialEquation(editor, target))

  const viewOnly = readOnly || initial.locked
  const { latex, setLatex, textarea, insertTemplate } = useEquationSource(initial.latex, viewOnly)
  const [display, setDisplay] = useState(initial.display)
  const result = useMemo(() => (latex.trim() === '' ? null : latexToMathMl(latex, display)), [latex, display])
  const shown = initial.locked ? initial.mathml : result?.ok === true ? result.tree : null

  function close(): void {
    onClose()
    // After the dialog leaves, so focus does not land on the window body.
    requestAnimationFrame(() => editor.view.focus())
  }

  function commit(): void {
    if (viewOnly || result === null || !result.ok) return
    const unchanged = target.kind === 'edit' && latex === initial.latex && display === initial.display
    if (!unchanged) {
      const content = { latex, mathml: result.mathml, display }
      if (target.kind === 'insert') insertEquation(editor, content)
      else replaceEquation(editor, target.pos, content)
    }
    close()
  }

  const error = result !== null && !result.ok ? result.error : null

  return (
    <div
      className="popover popover--wide equation"
      role="dialog"
      aria-label={t('document.math.dialog.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.preventDefault()
          close()
        } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
          event.preventDefault()
          commit()
        }
      }}
    >
      <EquationNotice locked={initial.locked} lossy={initial.lossy} readOnly={readOnly} />

      {!initial.locked && (
        <label className="popover__field">
          <span>{t('document.math.dialog.latex')}</span>
          <textarea
            ref={textarea}
            className="equation__source"
            value={latex}
            readOnly={readOnly}
            spellCheck={false}
            rows={3}
            onChange={(event) => setLatex(event.target.value)}
          />
        </label>
      )}

      <EquationPreview shown={shown} error={error} />

      {!viewOnly && <EquationTools display={display} onDisplay={setDisplay} onInsert={insertTemplate} />}

      <EquationActions
        viewOnly={viewOnly}
        canCommit={result !== null && result.ok}
        onClose={close}
        onCommit={commit}
      />
    </div>
  )
}

interface InitialEquation {
  readonly latex: string
  readonly display: boolean
  readonly lossy: readonly string[]
  readonly locked: boolean
  readonly mathml: MathElement | null
}

function initialEquation(editor: Editor, target: EquationTarget): InitialEquation {
  if (target.kind === 'insert') {
    return {
      latex: '',
      display: target.display,
      lossy: [] as string[],
      locked: false,
      mathml: null as MathElement | null,
    }
  }
  const node = editor.state.doc.nodeAt(target.pos)
  const lossy = (node?.attrs['lossy'] as string[] | undefined) ?? []
  const tree = sanitizeMathMl(String(node?.attrs['mathml'] ?? ''))
  const stored = String(node?.attrs['latex'] ?? '')
  return {
    latex: stored !== '' ? stored : tree === null ? '' : mathMlToLatex(tree),
    display: node?.attrs['display'] === true,
    lossy,
    locked: node?.attrs['editable'] === false || lossy.length > 0,
    mathml: tree,
  }
}

/** Node by node, as in the document: no innerHTML. */
function useMathPreview(shown: MathElement | null): RefObject<HTMLDivElement | null> {
  const preview = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    const host = preview.current
    if (host === null) return
    host.replaceChildren(...(shown === null ? [] : [buildMath(shown, document)]))
  }, [shown])
  return preview
}

interface EquationSource {
  readonly latex: string
  readonly setLatex: (latex: string) => void
  readonly textarea: RefObject<HTMLTextAreaElement | null>
  /** The template enters the selection, and the cursor lands where one writes next. */
  readonly insertTemplate: (template: MathTemplate) => void
}

function useEquationSource(initial: string, viewOnly: boolean): EquationSource {
  const [latex, setLatex] = useState(initial)
  const textarea = useRef<HTMLTextAreaElement>(null)
  const pendingCaret = useRef<number | null>(null)

  useLayoutEffect(() => {
    const caret = pendingCaret.current
    if (caret === null || textarea.current === null) return
    pendingCaret.current = null
    textarea.current.focus()
    textarea.current.setSelectionRange(caret, caret)
  }, [latex])

  useEffect(() => {
    if (!viewOnly) textarea.current?.focus()
  }, [viewOnly])

  function insertTemplate(template: MathTemplate): void {
    const area = textarea.current
    const start = area?.selectionStart ?? latex.length
    const end = area?.selectionEnd ?? latex.length
    pendingCaret.current = start + template.caret
    setLatex(latex.slice(0, start) + template.latex + latex.slice(end))
  }

  return { latex, setLatex, textarea, insertTemplate }
}

function MathPalette({ onInsert }: { onInsert: (template: MathTemplate) => void }): React.JSX.Element {
  const t = useT()
  return (
    <div className="chars equation__palette" aria-label={t('document.math.dialog.palette')} role="group">
      {MATH_PALETTE.map((group) => {
        const groupLabel = t(group.labelKey)
        return (
          <section key={group.labelKey} className="chars__group">
            <h3 className="chars__label">{groupLabel}</h3>
            <div className="chars__grid equation__grid" role="group" aria-label={groupLabel}>
              {group.templates.map((template) => {
                const name = t(template.nameKey)
                return (
                  <button
                    key={template.latex}
                    type="button"
                    className="chars__char equation__template"
                    aria-label={name}
                    title={`${name} — ${template.latex.trim()}`}
                    onClick={() => onInsert(template)}
                  >
                    {template.label}
                  </button>
                )
              })}
            </div>
          </section>
        )
      })}
    </div>
  )
}

function EquationActions({
  viewOnly,
  canCommit,
  onClose,
  onCommit,
}: {
  viewOnly: boolean
  canCommit: boolean
  onClose: () => void
  onCommit: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__actions">
      <span className="popover__spacer" />
      {viewOnly ? (
        <button type="button" className="btn btn--primary" onClick={onClose}>
          {t('document.common.close')}
        </button>
      ) : (
        <>
          <button type="button" className="btn" onClick={onClose}>
            {t('document.common.cancel')}
          </button>
          <button type="button" className="btn btn--primary" disabled={!canCommit} onClick={onCommit}>
            {t('document.math.dialog.ok')}
          </button>
        </>
      )}
    </div>
  )
}

function EquationNotice({
  locked,
  lossy,
  readOnly,
}: {
  locked: boolean
  lossy: readonly string[]
  readOnly: boolean
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      {locked && (
        <p className="popover__hint">{t('document.math.dialog.locked', { constructs: lossy.join(', ') })}</p>
      )}
      {readOnly && !locked && <p className="popover__hint">{t('document.math.dialog.readOnly')}</p>}
    </>
  )
}

function EquationPreview({
  shown,
  error,
}: {
  shown: MathElement | null
  error: string | null
}): React.JSX.Element {
  const t = useT()
  const preview = useMathPreview(shown)
  return (
    <>
      <div className="equation__preview" aria-label={t('document.math.dialog.preview')} role="img">
        <div ref={preview} />
        {shown === null && error === null && (
          <span className="popover__hint">{t('document.math.dialog.empty')}</span>
        )}
      </div>
      {error !== null && (
        <p className="popover__error" role="alert">
          {t('document.math.dialog.error', { message: error })}
        </p>
      )}
    </>
  )
}

function EquationTools({
  display,
  onDisplay,
  onInsert,
}: {
  display: boolean
  onDisplay: (display: boolean) => void
  onInsert: (template: MathTemplate) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <>
      <label className="popover__check">
        <input type="checkbox" checked={display} onChange={(event) => onDisplay(event.target.checked)} />
        {t('document.math.dialog.display')}
      </label>

      <MathPalette onInsert={onInsert} />
    </>
  )
}
