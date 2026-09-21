import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  LineSpacingKind,
  MAX_LINE_FACTOR,
  MAX_SPACING_PT,
  MIN_LINE_FACTOR,
  TextAlignment,
  isValidParagraphDraft,
  type ParagraphDraft,
} from '@services/document/paragraph-format.js'
import { paragraphDraftAt } from '../extensions/paragraph-commands.js'
import { ParagraphIndentFields } from './ParagraphIndentFields.js'
import { ParagraphSpacingFields } from './ParagraphSpacingFields.js'
import { isCustomLineSpacing, lineSpacingChoice, lineSpacingFrom } from './paragraph-draft.js'

/**
 * Diálogo de parágrafo — o equivalente ao do Word e ao do Writer.
 *
 * Mesmo desenho da configuração de página (`PageSetupPanel`): um `popover` com
 * campos em linha, rascunho local, `Escape` fecha e `Aplicar` só funciona quando
 * o formulário está válido. Duas telas que fazem a mesma coisa de dois jeitos
 * diferentes custam mais a quem usa do que a quem escreve.
 *
 * Abre com o que o parágrafo do cursor **já** tem, inclusive o que veio do
 * arquivo: é a única forma de o diálogo ser também um jeito de **ler** a
 * formatação de um documento alheio.
 */
export function ParagraphDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const [draft, setDraft] = useState<ParagraphDraft>(() => paragraphDraftAt(editor))

  const valid = isValidParagraphDraft(draft)
  const change = <K extends keyof ParagraphDraft>(key: K, value: ParagraphDraft[K]): void =>
    setDraft({ ...draft, [key]: value })

  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  function apply(): void {
    if (!valid) return
    editor.chain().focus().setParagraphFormat(draft).run()
    onClose()

    // E de novo depois do fechamento. O `focus()` da cadeia já rodou, mas o
    // painel sai da tela em seguida, e o que ainda estiver com o foco do
    // documento sai com ele — o foco cai no corpo da página. Sem isto, quem
    // clica em "Aplicar" precisa clicar no texto antes de voltar a escrever.
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label="Parágrafo"
      // No elemento, e não numa escuta global: o painel de localizar faz igual, e
      // é o que permite fechar sem tirar o foco de quem está preenchendo.
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field">
          <span>Alinhamento</span>
          <select
            aria-label="Alinhamento"
            value={draft.align}
            autoFocus
            onChange={(event) => change('align', event.target.value as TextAlignment)}
          >
            <option value={TextAlignment.Left}>À esquerda</option>
            <option value={TextAlignment.Center}>Centralizado</option>
            <option value={TextAlignment.Right}>À direita</option>
            <option value={TextAlignment.Justify}>Justificado</option>
          </select>
        </label>

        <label className="popover__field">
          <span>Entrelinha</span>
          <select
            aria-label="Entrelinha"
            value={lineSpacingChoice(draft)}
            onChange={(event) => setDraft({ ...draft, ...lineSpacingFrom(event.target.value) })}
          >
            <option value="single">Simples</option>
            <option value="1.15">1,15</option>
            <option value="1.5">1,5</option>
            <option value="2">Duplo</option>
            <option value="multiple">Múltiplo…</option>
            <option value="at-least">Pelo menos (pt)</option>
          </select>
        </label>

        {/* O campo da medida só aparece quando a escolha pede número: um campo
            desabilitado ao lado de "Simples" só faria a pessoa clicar nele. */}
        {isCustomLineSpacing(draft) && (
          <label className="popover__field popover__field--narrow">
            <span>{draft.lineSpacingKind === LineSpacingKind.AtLeast ? 'Pontos' : 'Fator'}</span>
            <input
              type="number"
              min={draft.lineSpacingKind === LineSpacingKind.AtLeast ? 1 : MIN_LINE_FACTOR}
              max={draft.lineSpacingKind === LineSpacingKind.AtLeast ? MAX_SPACING_PT : MAX_LINE_FACTOR}
              step={draft.lineSpacingKind === LineSpacingKind.AtLeast ? 1 : 0.05}
              value={draft.lineSpacingValue}
              onChange={(event) => change('lineSpacingValue', Number(event.target.value))}
            />
          </label>
        )}
      </div>

      <ParagraphSpacingFields draft={draft} onChange={change} />

      <ParagraphIndentFields draft={draft} onChange={change} />

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.keepNext}
          onChange={(event) => change('keepNext', event.target.checked)}
        />
        <span>Manter com o próximo parágrafo</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid
          ? 'Vale para os parágrafos que a seleção tocar.'
          : 'Há campo vazio ou fora da faixa — o documento não aceitaria a medida.'}
      </p>

      <div className="popover__actions">
        <button
          type="button"
          className="btn"
          onClick={() => setDraft({ ...DEFAULT_PARAGRAPH_DRAFT, align: draft.align })}
        >
          Restaurar padrão
        </button>
        <span className="popover__spacer" />
        {/*
          `preventDefault` no `mousedown`: sem ele o botão toma o foco do
          documento no clique, e desmontá-lo junto com o painel o devolve ao
          corpo da página — depois do `focus()` da cadeia, que já rodou. Quem
          clicava em "Aplicar" tinha de clicar no texto antes de continuar
          escrevendo. Com o foco parado no editor, não há corrida nenhuma.
        */}
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          Cancelar
        </button>
        <button
          type="button"
          className="btn btn--primary"
          onMouseDown={keepFocus}
          onClick={apply}
          disabled={!valid}
        >
          Aplicar
        </button>
      </div>
    </div>
  )
}
