import { useRef, useState } from 'react'
import type { MessageKey } from '@shared/i18n/index.js'
import {
  DEFAULT_PAGE_SETUP,
  PAGE_NUMBER_FORMATS,
  PageOrientation,
  PageSize,
  contentWidthMm,
  isValidMargins,
  pageDimensionsMm,
  type Margins,
  type PageSetup,
} from '@services/document/model.js'
import { MIN_MARGIN_FOR_HEADER_MM, marginFitsHeaderOrFooter } from '@services/pdf/page-setup.js'
import { usesEvenAndOdd, usesTitlePage } from '@services/document/band.js'
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
  // O último campo de texto que teve o cursor, e onde: é nele que "Número da
  // página" e "Total de páginas" entram, como no Word.
  const lastField = useRef<{ field: 'header' | 'footer'; at: number }>({ field: 'footer', at: -1 })

  function remember(field: 'header' | 'footer', input: HTMLInputElement): void {
    lastField.current = { field, at: input.selectionStart ?? input.value.length }
  }

  function insertField(token: string): void {
    const { field, at } = lastField.current
    const text = draft[field]
    const where = at < 0 || at > text.length ? text.length : at
    setDraft({ ...draft, [field]: text.slice(0, where) + token + text.slice(where) })
    lastField.current = { field, at: where + token.length }
  }

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
            onSelect={(event) => remember('header', event.currentTarget)}
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
            onSelect={(event) => remember('footer', event.currentTarget)}
          />
        </label>

        <div className="popover__row">
          <button type="button" className="btn" onClick={() => insertField('{n}')}>
            {t('document.pageSetup.insertPageNumber')}
          </button>
          <button type="button" className="btn" onClick={() => insertField('{total}')}>
            {t('document.pageSetup.insertTotalPages')}
          </button>
        </div>

        <p className="popover__hint">{t('document.pageSetup.hint', { n: '{n}', total: '{total}' })}</p>

        <label className="popover__check">
          <input
            type="checkbox"
            checked={usesTitlePage(draft)}
            onChange={(event) => setDraft({ ...draft, titlePage: event.target.checked })}
          />
          {t('document.pageSetup.titlePage')}
        </label>
        <label className="popover__check">
          <input
            type="checkbox"
            checked={usesEvenAndOdd(draft)}
            onChange={(event) => setDraft({ ...draft, evenAndOddHeaders: event.target.checked })}
          />
          {t('document.pageSetup.evenAndOdd')}
        </label>

        {/* O Chromium desenha cabeçalho e rodapé dentro da margem e recorta o
            excedente: com margem apertada eles somem sem explicação. */}
        {needsRoomWarning && (
          <p className="popover__error">
            {t('document.pageSetup.marginWarning', { min: MIN_MARGIN_FOR_HEADER_MM })}
          </p>
        )}
      </fieldset>

      <fieldset className="popover__fieldset">
        <legend>{t('document.pageSetup.pageNumbering')}</legend>
        <div className="popover__row">
          <label className="popover__field">
            <span>{t('document.pageSetup.pageNumberFormat')}</span>
            <select
              value={draft.pageNumberFormat ?? 'decimal'}
              onChange={(event) =>
                setDraft({ ...draft, pageNumberFormat: event.target.value as PageSetup['pageNumberFormat'] })
              }
            >
              {PAGE_NUMBER_FORMATS.map((format) => (
                <option key={format} value={format}>
                  {t(`document.pageSetup.pageFormat.${format}`)}
                </option>
              ))}
            </select>
          </label>
          <label className="popover__field popover__field--narrow">
            <span>{t('document.pageSetup.pageNumberStart')}</span>
            <input
              type="number"
              min={0}
              value={draft.pageNumberStart ?? ''}
              placeholder="1"
              onChange={(event) => {
                const value = event.target.value.trim()
                const number = Math.round(Number(value))
                setDraft({
                  ...draft,
                  pageNumberStart: value === '' || !Number.isFinite(number) ? null : Math.max(0, number),
                })
              }}
            />
          </label>
        </div>
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
