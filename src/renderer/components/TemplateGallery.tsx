import { useEffect, useState } from 'react'
import type { TemplateEntry } from '@shared/types.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'

interface Listing {
  readonly builtin: readonly TemplateEntry[]
  readonly user: readonly TemplateEntry[]
}

/**
 * Arquivo → Novo a partir de modelo….
 *
 * Os modelos que vêm com o aplicativo e os da pasta do usuário, cada um com nome
 * e descrição — sem miniatura, que pediria renderizar o pacote só para mostrar.
 * "Criar" (ou o duplo clique) abre o documento novo; "Procurar…" aceita qualquer
 * `.dotx`; "Abrir pasta de modelos" mostra a pasta onde o modelo salvo passa a
 * aparecer aqui.
 */
export function TemplateGallery(): React.JSX.Element {
  const t = useT()
  const close = useWorkspace((state) => state.setTemplateGallery)
  const newFromTemplate = useWorkspace((state) => state.newFromTemplate)
  const showError = useWorkspace((state) => state.showError)
  const [listing, setListing] = useState<Listing | null>(null)
  const [selected, setSelected] = useState<TemplateEntry | null>(null)

  useEffect(() => {
    let alive = true
    void window.api.template.list({}).then(
      (result) => {
        if (!alive) return
        if (!result.ok) {
          showError(result.error)
          return
        }
        setListing(result.data)
        setSelected(result.data.builtin[0] ?? null)
      },
      () => undefined,
    )
    return () => {
      alive = false
    }
  }, [showError])

  async function create(template: TemplateEntry | null): Promise<void> {
    // A galeria fecha antes: o aviso de alterações não salvas e o diálogo de
    // procurar ficam por cima dela, e ela não tem mais o que fazer depois.
    close(false)
    await newFromTemplate(template === null ? null : { source: template.source, id: template.id })
  }

  function item(entry: TemplateEntry): React.JSX.Element {
    const active = selected?.source === entry.source && selected.id === entry.id
    return (
      <li key={`${entry.source}:${entry.id}`}>
        <button
          type="button"
          role="option"
          aria-selected={active}
          className={active ? 'templates__item templates__item--active' : 'templates__item'}
          onClick={() => setSelected(entry)}
          onDoubleClick={() => void create(entry)}
        >
          <span className="templates__name">{entry.name}</span>
          <span className="templates__description">{entry.description}</span>
        </button>
      </li>
    )
  }

  return (
    <div className="templates__backdrop">
      <div
        className="popover popover--wide templates"
        role="dialog"
        aria-modal="true"
        aria-label={t('shell.template.title')}
        onKeyDown={(event) => {
          if (event.key === 'Escape') close(false)
          if (event.key === 'Enter' && selected !== null) void create(selected)
        }}
      >
        <h2 className="templates__title">{t('shell.template.title')}</h2>
        {listing === null ? (
          <p className="templates__empty">{t('shell.template.loading')}</p>
        ) : (
          <div className="templates__lists">
            <section>
              <h3>{t('shell.template.builtin')}</h3>
              <ul role="listbox" aria-label={t('shell.template.builtin')}>
                {listing.builtin.map(item)}
              </ul>
            </section>
            <section>
              <h3>{t('shell.template.user')}</h3>
              {listing.user.length === 0 ? (
                <p className="templates__empty">{t('shell.template.userEmpty')}</p>
              ) : (
                <ul role="listbox" aria-label={t('shell.template.user')}>
                  {listing.user.map(item)}
                </ul>
              )}
            </section>
          </div>
        )}

        <div className="popover__actions">
          <button type="button" className="btn" onClick={() => void create(null)}>
            {t('shell.template.browse')}
          </button>
          <button type="button" className="btn" onClick={() => void window.api.template.openFolder({})}>
            {t('shell.template.openFolder')}
          </button>
          <span className="popover__spacer" />
          <button type="button" className="btn" onClick={() => close(false)}>
            {t('shell.template.cancel')}
          </button>
          <button
            type="button"
            className="btn btn--primary"
            autoFocus
            disabled={selected === null}
            onClick={() => void create(selected)}
          >
            {t('shell.template.create')}
          </button>
        </div>
      </div>
    </div>
  )
}
