import { useState } from 'react'
import {
  DEFAULT_PAGE_SETUP,
  PageOrientation,
  PageSize,
  contentWidthMm,
  isValidMargins,
  pageDimensionsMm,
  type Margins,
  type PageSetup,
} from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'

const MARGIN_FIELDS: readonly { readonly key: keyof Margins; readonly label: string }[] = [
  { key: 'top', label: 'Superior' },
  { key: 'bottom', label: 'Inferior' },
  { key: 'left', label: 'Esquerda' },
  { key: 'right', label: 'Direita' },
]

export function PageSetupPanel({ onClose }: { readonly onClose: () => void }): React.JSX.Element {
  const current = useWorkspace((state) => state.page)
  const setPage = useWorkspace((state) => state.setPage)
  const [draft, setDraft] = useState<PageSetup>(current)

  const valid = isValidMargins(draft)
  const { width, height } = pageDimensionsMm(draft)

  function apply(): void {
    if (!valid) return
    setPage(draft)
    onClose()
  }

  return (
    <div className="popover popover--wide" role="dialog" aria-label="Configuração de página">
      <div className="popover__row">
        <label className="popover__field">
          <span>Tamanho</span>
          <select
            value={draft.size}
            onChange={(event) => setDraft({ ...draft, size: event.target.value as PageSize })}
          >
            <option value={PageSize.A4}>A4 (210 × 297 mm)</option>
            <option value={PageSize.Letter}>Carta (216 × 279 mm)</option>
          </select>
        </label>

        <label className="popover__field">
          <span>Orientação</span>
          <select
            value={draft.orientation}
            onChange={(event) => setDraft({ ...draft, orientation: event.target.value as PageOrientation })}
          >
            <option value={PageOrientation.Portrait}>Retrato</option>
            <option value={PageOrientation.Landscape}>Paisagem</option>
          </select>
        </label>
      </div>

      <fieldset className="popover__fieldset">
        <legend>Margens (mm)</legend>
        <div className="popover__row">
          {MARGIN_FIELDS.map(({ key, label }) => (
            <label key={key} className="popover__field popover__field--narrow">
              <span>{label}</span>
              <input
                type="number"
                min={0}
                max={100}
                step={1}
                value={draft.margins[key]}
                onChange={(event) =>
                  setDraft({
                    ...draft,
                    margins: { ...draft.margins, [key]: Number(event.target.value) },
                  })
                }
              />
            </label>
          ))}
        </div>
      </fieldset>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid
          ? `Página ${width} × ${height} mm · área de texto ${contentWidthMm(draft).toFixed(0)} mm de largura`
          : 'As margens somam mais que a página e não deixam espaço para o texto.'}
      </p>

      <div className="popover__actions">
        <button type="button" className="btn" onClick={() => setDraft(DEFAULT_PAGE_SETUP)}>
          Restaurar padrão
        </button>
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          Cancelar
        </button>
        <button type="button" className="btn btn--primary" onClick={apply} disabled={!valid}>
          Aplicar
        </button>
      </div>
    </div>
  )
}
