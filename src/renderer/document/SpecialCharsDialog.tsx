import { useEffect, useRef } from 'react'
import type { Editor } from '@tiptap/react'
import { SPECIAL_CHARACTER_GROUPS, type SpecialCharacter } from '@services/document/special-characters.js'
import { useT } from '../i18n.js'
import { cycleFocus } from '../components/focus-trap.js'

/**
 * O foco **entra no painel** ao abrir, e o `Tab` **fica dentro** dele: senão as
 * setas andariam pelo documento, e o `Tab` cairia no seletor "Estilo" da barra,
 * onde a seta seguinte trocaria o estilo do parágrafo.
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

  // No primeiro caractere, e não no "Fechar": quem abriu quer inserir.
  useEffect(() => {
    host.current?.querySelector<HTMLButtonElement>('button[data-char]')?.focus()
  }, [])

  function insert(character: SpecialCharacter): void {
    // Insere e continua no painel: quem abre costuma querer dois ou três
    // caracteres. Sem foco no editor a inserção cai no lugar certo, porque a
    // seleção do ProseMirror avança a cada uma.
    editor.chain().insertContent(character.char).run()
  }

  /** No elemento de fora, para o `Escape` valer com o foco no "Fechar". */
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
                      // Um botão chamado "—" não diz nada a quem não o vê.
                      aria-label={charName}
                      title={charName}
                      onClick={() => insert(character)}
                    >
                      {/* O espaço inquebrável não desenha nada: sem marcador, o botão pareceria quebrado. */}
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

/** A conta sai da posição desenhada: a grade é fluida. */
function columnsAround(buttons: readonly HTMLButtonElement[], current: number): number {
  const reference = buttons[current]
  if (reference === undefined) return 1

  const top = reference.offsetTop
  const row = buttons.filter((button) => button.offsetTop === top)
  return Math.max(1, row.length)
}

/** O passo vertical é o número de colunas medido no DOM: a grade é fluida. */
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
