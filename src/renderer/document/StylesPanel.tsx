import { useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import {
  StyleType,
  blockStyleOf,
  listedStyles,
  styleLabelOf,
  type StyleDefinition,
} from '@services/document/styles.js'
import { useWorkspace } from '../state/workspace.js'

/**
 * Os estilos do documento, só de leitura.
 *
 * Responde a duas perguntas que até aqui não tinham resposta nenhuma na tela:
 * **quais estilos este documento tem** e **qual é o do parágrafo onde estou**. Um
 * `.docx` corporativo guarda quase toda a formatação em estilos, e quem abria um
 * aqui não tinha como saber disso — a tela mostrava o resultado, nunca a regra.
 *
 * Só de leitura de propósito, e não por falta de tempo: aplicar um estilo é
 * reescrever o `w:pStyle` de um parágrafo, e modificá-lo é reescrever
 * `word/styles.xml` — as duas coisas mexem no arquivo de quem confia neste
 * programa, e cada uma tem a sua entrega. Enquanto isso não existe, é melhor uma
 * tela que mostra do que um botão que promete.
 *
 * Mesmo desenho dos outros painéis (`PageSetupPanel`, `ParagraphDialog`): um
 * `popover`, `Escape` fecha, e o foco começa no botão de fechar.
 */
export function StylesPanel({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const sheet = useWorkspace((state) => state.styles)
  const [filter, setFilter] = useState<'all' | StyleType>('all')

  // Ao vivo: o cursor anda enquanto o painel está aberto, e o estilo do bloco
  // muda com ele. Fechar o painel não é condição para continuar escrevendo.
  const block = useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const parent = current.state.selection.$from.parent
      const styleId = parent.attrs['styleId']
      const level = parent.attrs['level']
      return {
        type: parent.type.name,
        styleId: typeof styleId === 'string' ? styleId : null,
        level: typeof level === 'number' ? level : null,
      }
    },
  })

  const current = blockStyleOf(sheet, block)
  const styles = listedStyles(sheet).filter((style) => filter === 'all' || style.type === filter)

  return (
    <div
      className="popover"
      role="dialog"
      aria-label="Estilos"
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <p className="popover__hint">
        {current === null
          ? 'Este documento não define estilos para o parágrafo do cursor.'
          : `Parágrafo do cursor: ${styleLabelOf(current)}`}
      </p>

      <div className="popover__row">
        <label className="popover__field">
          <span>Mostrar</span>
          <select
            aria-label="Mostrar"
            value={filter}
            onChange={(event) => setFilter(event.target.value as 'all' | StyleType)}
          >
            <option value="all">Todos os estilos</option>
            <option value={StyleType.Paragraph}>De parágrafo</option>
            <option value={StyleType.Character}>De caractere</option>
          </select>
        </label>
      </div>

      {/* Focalizável para que a lista role pelo teclado: não há o que escolher
          aqui, então um `listbox` prometeria uma seleção que não existe. */}
      <ul className="styles-list" tabIndex={0} aria-label="Estilos do documento">
        {styles.length === 0 && <li className="styles-list__empty">Nenhum estilo deste tipo.</li>}
        {styles.map((style) => (
          <li
            key={style.id}
            className={
              style.id === current?.id ? 'styles-list__item styles-list__item--current' : 'styles-list__item'
            }
            // O leitor de tela precisa ouvir "este é o do cursor" junto com o
            // nome; a marca visual sozinha não diz nada a quem não vê a tela.
            aria-current={style.id === current?.id ? 'true' : undefined}
          >
            <span className="styles-list__name">{styleLabelOf(style)}</span>
            <span className="styles-list__meta">{describe(style)}</span>
          </li>
        ))}
      </ul>

      <p className="popover__hint">
        Por enquanto o painel apenas mostra: aplicar, criar e modificar estilos vêm nas próximas versões.
      </p>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn btn--primary" autoFocus onClick={onClose}>
          Fechar
        </button>
      </div>
    </div>
  )
}

/**
 * O resumo de um estilo numa linha.
 *
 * A fonte e o tamanho, quando o estilo os declara, mais de quem ele herda: é o
 * que responde "por que este parágrafo está assim" sem abrir o arquivo. O que o
 * estilo não declara não aparece — dizer "12 pt" num estilo que herda o tamanho
 * seria afirmar algo que o documento não diz, e a cascata é da entrega seguinte.
 */
function describe(style: StyleDefinition): string {
  const parts: string[] = [style.type === StyleType.Character ? 'caractere' : 'parágrafo']

  const font = style.character?.fontFamily
  if (font !== undefined) parts.push(font.split(',')[0]!.trim())
  if (style.character?.fontSize !== undefined) parts.push(style.character.fontSize)
  if (style.character?.bold === true) parts.push('negrito')
  if (style.basedOn !== undefined) parts.push(`baseado em ${style.basedOn}`)

  return parts.join(' · ')
}
