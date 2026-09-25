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
import { useT } from '../../i18n.js'
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
  const t = useT()
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
      aria-label={t('document.paragraph.title')}
      // No elemento, e não numa escuta global: o painel de localizar faz igual, e
      // é o que permite fechar sem tirar o foco de quem está preenchendo.
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('document.paragraph.alignment')}</span>
          <select
            aria-label={t('document.paragraph.alignment')}
            value={draft.align}
            autoFocus
            onChange={(event) => change('align', event.target.value as TextAlignment)}
          >
            <option value={TextAlignment.Left}>{t('document.paragraph.alignLeft')}</option>
            <option value={TextAlignment.Center}>{t('document.paragraph.alignCenter')}</option>
            <option value={TextAlignment.Right}>{t('document.paragraph.alignRight')}</option>
            <option value={TextAlignment.Justify}>{t('document.paragraph.alignJustify')}</option>
          </select>
        </label>

        <label className="popover__field">
          <span>{t('document.paragraph.lineSpacing')}</span>
          <select
            aria-label={t('document.paragraph.lineSpacing')}
            value={lineSpacingChoice(draft)}
            onChange={(event) => setDraft({ ...draft, ...lineSpacingFrom(event.target.value) })}
          >
            <option value="single">{t('document.paragraph.spacingSingle')}</option>
            <option value="1.15">1,15</option>
            <option value="1.5">1,5</option>
            <option value="2">{t('document.paragraph.spacingDouble')}</option>
            <option value="multiple">{t('document.paragraph.spacingMultiple')}</option>
            <option value="at-least">{t('document.paragraph.spacingAtLeast')}</option>
          </select>
        </label>

        {/* O campo da medida só aparece quando a escolha pede número: um campo
            desabilitado ao lado de "Simples" só faria a pessoa clicar nele. */}
        {isCustomLineSpacing(draft) && (
          <label className="popover__field popover__field--narrow">
            <span>
              {draft.lineSpacingKind === LineSpacingKind.AtLeast
                ? t('document.paragraph.points')
                : t('document.paragraph.factor')}
            </span>
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
        <span>{t('document.paragraph.keepWithNext')}</span>
      </label>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.keepLines}
          onChange={(event) => change('keepLines', event.target.checked)}
        />
        <span>{t('document.paragraph.keepLinesTogether')}</span>
      </label>

      <label className="popover__check">
        <input
          type="checkbox"
          checked={draft.widowControl}
          onChange={(event) => change('widowControl', event.target.checked)}
        />
        <span>{t('document.paragraph.widowControl')}</span>
      </label>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid ? t('document.paragraph.hintValid') : t('document.paragraph.hintInvalid')}
      </p>

      <div className="popover__actions">
        <button
          type="button"
          className="btn"
          onClick={() => setDraft({ ...DEFAULT_PARAGRAPH_DRAFT, align: draft.align })}
        >
          {t('document.common.restoreDefaults')}
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
          {t('document.common.cancel')}
        </button>
        <button
          type="button"
          className="btn btn--primary"
          onMouseDown={keepFocus}
          onClick={apply}
          disabled={!valid}
        >
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
