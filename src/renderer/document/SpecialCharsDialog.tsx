import { useEffect, useRef } from 'react'
import type { Editor } from '@tiptap/react'
import { SPECIAL_CHARACTER_GROUPS, type SpecialCharacter } from '@services/document/special-characters.js'

/**
 * Seletor de caracteres especiais.
 *
 * O conjunto que o Word deixa à mão, agrupado pelo que se procura: pontuação,
 * moeda, matemática, grego e marcas. O catálogo mora em `@services` porque é
 * dado, não tela.
 *
 * **Navegável pelo teclado sem truque**: cada caractere é um botão de verdade, e
 * as setas andam pela grade como se espera de uma grade.
 *
 * Duas decisões fazem o teclado funcionar de verdade, e as duas faltavam:
 *
 *  - o **foco entra no painel** ao abrir. Sem isso o foco continuava no texto: as
 *    setas moviam o cursor do documento, `Enter` partia o parágrafo e só o mouse
 *    inseria;
 *  - o `Tab` **fica dentro** do painel. Sem a armadilha, o `Tab` caía no seletor
 *    "Estilo" da barra de ferramentas, e a seta seguinte trocava o estilo do
 *    parágrafo — um painel de inserir símbolo mudava o documento.
 */
export function SpecialCharsDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const panel = useRef<HTMLDivElement>(null)
  const host = useRef<HTMLDivElement>(null)

  // O foco vai para o primeiro caractere da grade, e não para o botão "Fechar":
  // quem abriu o painel quer inserir, e daí as setas já andam pela grade.
  useEffect(() => {
    host.current?.querySelector<HTMLButtonElement>('button[data-char]')?.focus()
  }, [])

  function insert(character: SpecialCharacter): void {
    // Insere e **continua no painel** — sem `focus()` no editor. Duas razões: quem
    // abre este painel costuma querer dois ou três caracteres (fechar a cada um
    // seria um clique por símbolo), e devolver o foco ao texto faria a seta
    // seguinte andar pelo documento em vez de pela grade, que é exatamente o
    // defeito que o painel tinha.
    //
    // A inserção cai no lugar certo mesmo com o editor sem foco: o cursor mora na
    // seleção do ProseMirror, e cada inserção a empurra para depois do que
    // acabou de entrar.
    editor.chain().insertContent(character.char).run()
  }

  /**
   * As setas andam pela grade.
   *
   * O passo horizontal é um botão; o vertical é o número de colunas que couberam
   * na linha, medido no próprio DOM — a grade é fluida, e um número fixo aqui
   * discordaria do que está na tela em outra largura de janela.
   */
  function onKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void {
    const sideways = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0
    const updown = event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0
    if (sideways === 0 && updown === 0) return

    const buttons = [...(host.current?.querySelectorAll<HTMLButtonElement>('button[data-char]') ?? [])]
    const current = buttons.indexOf(event.target as HTMLButtonElement)
    if (current === -1) return

    const offset = sideways === 0 ? updown * columnsAround(buttons, current) : sideways
    const target = buttons[current + offset]
    if (target === undefined) return

    event.preventDefault()
    target.focus()
  }

  /**
   * `Escape` fecha, e `Tab` circula dentro do painel.
   *
   * No elemento de fora, e não na grade: o botão "Fechar" também é do painel, e
   * `Escape` tem de valer com o foco nele. A lista de paradas sai do DOM porque é
   * ela que muda de tamanho — a grade tem centenas de botões, e o último é o
   * "Fechar".
   */
  function onPanelKeyDown(event: React.KeyboardEvent<HTMLDivElement>): void {
    if (event.key === 'Escape') {
      onClose()
      return
    }

    if (event.key !== 'Tab') return

    const stops = [...(panel.current?.querySelectorAll<HTMLButtonElement>('button') ?? [])]
    if (stops.length === 0) return

    const current = stops.indexOf(event.target as HTMLButtonElement)
    const next = stops[(current + (event.shiftKey ? -1 : 1) + stops.length) % stops.length]
    if (next === undefined) return

    event.preventDefault()
    next.focus()
  }

  return (
    <div
      ref={panel}
      className="popover popover--wide"
      role="dialog"
      aria-label="Caracteres especiais"
      onKeyDown={onPanelKeyDown}
    >
      <div ref={host} className="chars" onKeyDown={onKeyDown}>
        {SPECIAL_CHARACTER_GROUPS.map((group) => (
          <section key={group.label} className="chars__group">
            <h3 className="chars__label">{group.label}</h3>
            <div className="chars__grid" role="group" aria-label={group.label}>
              {group.characters.map((character) => (
                <button
                  key={character.char}
                  type="button"
                  className="chars__char"
                  data-char={character.char}
                  // O nome acessível é o nome do caractere: um botão chamado "—"
                  // não diz nada a quem não o vê. A dica do mouse repete, para
                  // quem vê e não sabe o nome.
                  aria-label={character.name}
                  title={character.name}
                  onClick={() => insert(character)}
                >
                  {/* O espaço inquebrável não desenha nada: sem um marcador, o
                      botão dele pareceria vazio e quebrado. */}
                  {character.char === '\u00a0' ? '␣' : character.char}
                </button>
              ))}
            </div>
          </section>
        ))}
      </div>

      <div className="popover__actions">
        <span className="popover__hint">
          Setas escolhem, Enter insere. O painel continua aberto, e Esc fecha.
        </span>
        <span className="popover__spacer" />
        <button type="button" className="btn btn--primary" onClick={onClose}>
          Fechar
        </button>
      </div>
    </div>
  )
}

/**
 * Quantos botões cabem na linha em que este botão está.
 *
 * Para as setas horizontais o passo é sempre 1; para as verticais, é a largura da
 * linha. A conta sai da posição desenhada, e não de uma constante: a grade é
 * fluida.
 */
function columnsAround(buttons: readonly HTMLButtonElement[], current: number): number {
  const reference = buttons[current]
  if (reference === undefined) return 1

  const top = reference.offsetTop
  const row = buttons.filter((button) => button.offsetTop === top)
  return Math.max(1, row.length)
}
