import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { latexToMathMl } from '@services/document/latex.js'
import { mathMlToLatex } from '@services/document/mathml-latex.js'
import { sanitizeMathMl, type MathElement } from '@services/document/mathml.js'
import { MATH_PALETTE, type MathTemplate } from '@services/document/math-palette.js'
import { buildMath } from './extensions/math.js'
import { insertEquation, replaceEquation, type EquationTarget } from './math-commands.js'
import { useT } from '../i18n.js'

/**
 * O editor de equações (M11, fase 2): o LaTeX, a visualização ao vivo e os
 * modelos.
 *
 * O LaTeX é a fonte; o Temml o desenha em MathML (`latex.ts`), que passa pelo
 * mesmo filtro da equação lida do arquivo antes de virar DOM. A equação que veio
 * de um `.docx` não tem LaTeX: ele sai do MathML dela (`mathml-latex.ts`) ao
 * abrir. Se ninguém mudar nada, o OK fecha sem tocar no documento — e a equação
 * continua com o OMML do arquivo, intacto.
 *
 * A equação travada (com construções que a tela não desenha) abre só para ver:
 * editar o LaTeX derivado dela perderia justamente o que não se vê.
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

  // O nó lido uma vez, ao abrir: o diálogo edita o que estava lá naquele momento.
  const [initial] = useState(() => {
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
  })

  const [latex, setLatex] = useState(initial.latex)
  const [display, setDisplay] = useState(initial.display)
  const textarea = useRef<HTMLTextAreaElement>(null)
  const preview = useRef<HTMLDivElement>(null)
  const pendingCaret = useRef<number | null>(null)

  const viewOnly = readOnly || initial.locked
  const result = useMemo(() => (latex.trim() === '' ? null : latexToMathMl(latex, display)), [latex, display])
  const shown = initial.locked ? initial.mathml : result?.ok === true ? result.tree : null

  // A visualização é montada nó a nó, como a equação no documento: nada de innerHTML.
  useLayoutEffect(() => {
    const host = preview.current
    if (host === null) return
    host.replaceChildren(...(shown === null ? [] : [buildMath(shown, document)]))
  }, [shown])

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

  function close(): void {
    onClose()
    // Depois de o diálogo sair: o foco volta ao texto, e não ao corpo da janela.
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

  function insertTemplate(template: MathTemplate): void {
    const area = textarea.current
    const start = area?.selectionStart ?? latex.length
    const end = area?.selectionEnd ?? latex.length
    pendingCaret.current = start + template.caret
    setLatex(latex.slice(0, start) + template.latex + latex.slice(end))
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
      {initial.locked && (
        <p className="popover__hint">
          {t('document.math.dialog.locked', { constructs: initial.lossy.join(', ') })}
        </p>
      )}
      {readOnly && !initial.locked && <p className="popover__hint">{t('document.math.dialog.readOnly')}</p>}

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

      {!viewOnly && (
        <>
          <label className="popover__check">
            <input type="checkbox" checked={display} onChange={(event) => setDisplay(event.target.checked)} />
            {t('document.math.dialog.display')}
          </label>

          <div
            className="chars equation__palette"
            aria-label={t('document.math.dialog.palette')}
            role="group"
          >
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
                          onClick={() => insertTemplate(template)}
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
        </>
      )}

      <div className="popover__actions">
        <span className="popover__spacer" />
        {viewOnly ? (
          <button type="button" className="btn btn--primary" onClick={close}>
            {t('document.common.close')}
          </button>
        ) : (
          <>
            <button type="button" className="btn" onClick={close}>
              {t('document.common.cancel')}
            </button>
            <button
              type="button"
              className="btn btn--primary"
              disabled={result === null || !result.ok}
              onClick={commit}
            >
              {t('document.math.dialog.ok')}
            </button>
          </>
        )}
      </div>
    </div>
  )
}
