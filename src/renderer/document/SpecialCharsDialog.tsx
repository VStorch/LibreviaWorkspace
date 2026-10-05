import { useEffect, useRef } from 'react'
import type { Editor } from '@tiptap/react'
import { SPECIAL_CHARACTER_GROUPS, type SpecialCharacter } from '@services/document/special-characters.js'
import { useT } from '../i18n.js'
import { cycleFocus } from '../components/focus-trap.js'

/**
 * Focus **enters the panel** on open, and `Tab` **stays inside** it: otherwise the arrows would
 * move through the document, and `Tab` would land on the toolbar's "Style" picker, where the next
 * arrow would change the paragraph style.
 */
export function SpecialCharsDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const panel = useRef<HTMLDivElement>(null)
  const host = useRef<HTMLDivElement>(null)

  // On the first character, not on "Close": whoever opened it wants to insert.
  useEffect(() => {
    host.current?.querySelector<HTMLButtonElement>('button[data-char]')?.focus()
  }, [])

  function insert(character: SpecialCharacter): void {
    // Inserts and stays in the panel: people usually want two or three characters. Without focus on
    // the editor the insertion lands in the right place, because the ProseMirror selection advances
    // each time.
    editor.chain().insertContent(character.char).run()
  }

  /** On the outer element, so `Escape` works with focus on "Close". */
  function onPanelKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void {
    if (event.key === 'Escape') onClose()
    else if (event.key === 'Tab') cycleFocus(panel.current, event, 'button')
  }

  return (
    <div
      ref={panel}
      className="popover popover--wide"
      role="dialog"
      aria-label={t('document.specialChars.title')}
      onKeyDown={onPanelKeyDown}
    >
      <div ref={host} className="chars" onKeyDown={(event) => moveAmongCharacters(host.current, event)}>
        {SPECIAL_CHARACTER_GROUPS.map((group) => {
          const groupLabel = t(group.labelKey)
          return (
            <section key={group.labelKey} className="chars__group">
              <h3 className="chars__label">{groupLabel}</h3>
              <div className="chars__grid" role="group" aria-label={groupLabel}>
                {group.characters.map((character) => {
                  const charName = t(character.nameKey)
                  return (
                    <button
                      key={character.char}
                      type="button"
                      className="chars__char"
                      data-char={character.char}
                      // A button called "—" says nothing to someone who cannot see it.
                      aria-label={charName}
                      title={charName}
                      onClick={() => insert(character)}
                    >
                      {/* A non-breaking space draws nothing: without a marker, the button would look broken.
                       */}
                      {character.char === '\u00a0' ? '␣' : character.char}
                    </button>
                  )
                })}
              </div>
            </section>
          )
        })}
      </div>

      <div className="popover__actions">
        <span className="popover__hint">{t('document.specialChars.hint')}</span>
        <span className="popover__spacer" />
        <button type="button" className="btn btn--primary" onClick={onClose}>
          {t('document.common.close')}
        </button>
      </div>
    </div>
  )
}

/** Computed from the drawn position: the grid is fluid. */
function columnsAround(buttons: readonly HTMLButtonElement[], current: number): number {
  const reference = buttons[current]
  if (reference === undefined) return 1

  const top = reference.offsetTop
  const row = buttons.filter((button) => button.offsetTop === top)
  return Math.max(1, row.length)
}

/** The vertical step is the column count measured in the DOM: the grid is fluid. */
function moveAmongCharacters(host: HTMLDivElement | null, event: React.KeyboardEvent<HTMLDivElement>): void {
  const sideways = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0
  const updown = event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0
  if (sideways === 0 && updown === 0) return

  const buttons = [...(host?.querySelectorAll<HTMLButtonElement>('button[data-char]') ?? [])]
  const current = buttons.indexOf(event.target as HTMLButtonElement)
  if (current === -1) return

  const offset = sideways === 0 ? updown * columnsAround(buttons, current) : sideways
  const target = buttons[current + offset]
  if (target === undefined) return

  event.preventDefault()
  target.focus()
}
