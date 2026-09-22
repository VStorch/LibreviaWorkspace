import { useState } from 'react'
import type { MessageKey } from '@shared/i18n/index.js'
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
import { MIN_MARGIN_FOR_HEADER_MM, marginFitsHeaderOrFooter } from '@services/pdf/page-setup.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'

const MARGIN_FIELDS: readonly { readonly key: keyof Margins; readonly labelKey: MessageKey }[] = [
  { key: 'top', labelKey: 'document.pageSetup.marginTop' },
  { key: 'bottom', labelKey: 'document.pageSetup.marginBottom' },
  { key: 'left', labelKey: 'document.pageSetup.marginLeft' },
  { key: 'right', labelKey: 'document.pageSetup.marginRight' },
]

export function PageSetupPanel({ onClose }: { readonly onClose: () => void }): React.JSX.Element {
  const t = useT()
  const current = useWorkspace((state) => state.page)
  const setPage = useWorkspace((state) => state.setPage)
  const [draft, setDraft] = useState<PageSetup>(current)

  const valid = isValidMargins(draft)
  const { width, height } = pageDimensionsMm(draft)

  const usesHeaderOrFooter = draft.header.trim().length > 0 || draft.footer.trim().length > 0
  const needsRoomWarning =
    usesHeaderOrFooter &&
    (!marginFitsHeaderOrFooter(draft.margins.top) || !marginFitsHeaderOrFooter(draft.margins.bottom))

  function apply(): void {
    if (!valid) return
    setPage(draft)
    onClose()
  }

  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label={t('document.pageSetup.title')}
      // `Esc` fecha e `Enter` aplica, como no diálogo de parágrafo: dois painéis
      // que fazem a mesma coisa de dois jeitos custam mais a quem usa do que a
      // quem escreve. No elemento, e não numa escuta global, para não fechar
      // enquanto o foco está em outro canto da tela.
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('document.pageSetup.size')}</span>
          <select
            aria-label={t('document.pageSetup.size')}
            value={draft.size}
            onChange={(event) => setDraft({ ...draft, size: event.target.value as PageSize })}
          >
            <option value={PageSize.A4}>A4 (210 × 297 mm)</option>
            <option value={PageSize.Letter}>{t('document.pageSetup.letter')}</option>
          </select>
        </label>

        <label className="popover__field">
          <span>{t('document.pageSetup.orientation')}</span>
          <select
            aria-label={t('document.pageSetup.orientation')}
            value={draft.orientation}
            onChange={(event) => setDraft({ ...draft, orientation: event.target.value as PageOrientation })}
          >
            <option value={PageOrientation.Portrait}>{t('document.pageSetup.portrait')}</option>
            <option value={PageOrientation.Landscape}>{t('document.pageSetup.landscape')}</option>
          </select>
        </label>
      </div>

      <fieldset className="popover__fieldset">
        <legend>{t('document.pageSetup.margins')}</legend>
        <div className="popover__row">
          {MARGIN_FIELDS.map(({ key, labelKey }) => (
            <label key={key} className="popover__field popover__field--narrow">
              <span>{t(labelKey)}</span>
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

      <fieldset className="popover__fieldset">
        <legend>{t('document.pageSetup.headerAndFooter')}</legend>

        <label className="popover__field">
          <span>{t('document.pageSetup.header')}</span>
          <input
            type="text"
            value={draft.header}
            placeholder={t('document.pageSetup.headerPlaceholder')}
            maxLength={500}
            onChange={(event) => setDraft({ ...draft, header: event.target.value })}
          />
        </label>

        <label className="popover__field">
          <span>{t('document.pageSetup.footer')}</span>
          <input
            type="text"
            value={draft.footer}
            placeholder={t('document.pageSetup.footerPlaceholder')}
            maxLength={500}
            onChange={(event) => setDraft({ ...draft, footer: event.target.value })}
          />
        </label>

        <p className="popover__hint">
          {t('document.pageSetup.hint', { n: '{n}', total: '{total}' })}
        </p>

        {/* O Chromium desenha cabeçalho e rodapé dentro da margem e recorta o
            excedente: com margem apertada eles somem sem explicação. */}
        {needsRoomWarning && (
          <p className="popover__error">
            {t('document.pageSetup.marginWarning', { min: MIN_MARGIN_FOR_HEADER_MM })}
          </p>
        )}
      </fieldset>

      <p className={valid ? 'popover__hint' : 'popover__error'}>
        {valid
          ? t('document.pageSetup.pageSummary', {
              width,
              height,
              contentWidth: contentWidthMm(draft).toFixed(0),
            })
          : t('document.pageSetup.marginsError')}
      </p>

      <div className="popover__actions">
        <button type="button" className="btn" onClick={() => setDraft(DEFAULT_PAGE_SETUP)}>
          {t('document.common.restoreDefaults')}
        </button>
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={apply} disabled={!valid}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}
