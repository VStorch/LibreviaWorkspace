import { useState } from 'react'
import { AiProviderKind } from '@services/ai/provider.js'
import { useT } from '../i18n.js'
import { AiModelSelect } from '../document/AiModelSelect.js'
import { useAiProviders, type AiProviderInfo, type AiProviders } from '../document/useAiProviders.js'

export interface AiPanelChoice {
  readonly providers: AiProviders
  readonly provider: AiProviderInfo | undefined
  readonly chooseProvider: (id: string) => void
  readonly chooseModel: (model: string) => void
}

/** The default assistant until the person picks another in the panel. */
export function useAiPanelChoice(configuring: boolean): AiPanelChoice {
  const providers = useAiProviders(configuring)
  const { list, update } = providers
  const [providerId, setProviderId] = useState<string | null>(null)
  const provider =
    list?.providers.find((candidate) => candidate.id === (providerId ?? list.defaultProviderId)) ??
    list?.providers[0]
  const chooseModel = (model: string): void => {
    if (provider !== undefined) void update(window.api.ai.chooseModel({ providerId: provider.id, model }))
  }
  return { providers, provider, chooseProvider: setProviderId, chooseModel }
}

interface AiPanelFrameProps {
  readonly choice: AiPanelChoice
  readonly onClose: () => void
  readonly onConfigure: () => void
  readonly children: React.ReactNode
}

/** The side panel of both editors: title, close, and the setup when there is no assistant yet. */
export function AiPanelFrame({
  choice,
  onClose,
  onConfigure,
  children,
}: AiPanelFrameProps): React.JSX.Element {
  const t = useT()
  const { list } = choice.providers
  return (
    <aside className="ai-pane" aria-label={t('ai.panel.title')}>
      <div className="nav-pane__header">
        <span className="nav-pane__title">{t('ai.panel.title')}</span>
        <button
          type="button"
          className="nav-pane__close"
          aria-label={t('document.common.close')}
          title={t('document.common.close')}
          onClick={onClose}
        >
          ×
        </button>
      </div>
      {list !== null && choice.provider === undefined ? (
        <AiNoProviders onConfigure={onConfigure} />
      ) : (
        <div className="ai-pane__body">{children}</div>
      )}
    </aside>
  )
}

function AiNoProviders({ onConfigure }: { readonly onConfigure: () => void }): React.JSX.Element {
  const t = useT()
  return (
    <div className="ai-pane__body">
      <p className="ai-pane__note">{t('ai.panel.noProviders')}</p>
      <button type="button" className="btn btn--primary" onClick={onConfigure}>
        {t('ai.panel.configure')}
      </button>
    </div>
  )
}

export function AiProviderFields({
  choice,
  running,
}: {
  readonly choice: AiPanelChoice
  readonly running: boolean
}): React.JSX.Element {
  const t = useT()
  const { provider } = choice
  return (
    <>
      <label className="popover__field">
        <span>{t('ai.panel.provider')}</span>
        <select value={provider?.id ?? ''} onChange={(event) => choice.chooseProvider(event.target.value)}>
          {(choice.providers.list?.providers ?? []).map((candidate) => (
            <option key={candidate.id} value={candidate.id}>
              {candidate.label}
            </option>
          ))}
        </select>
      </label>
      {provider !== undefined && provider.kind !== AiProviderKind.ClaudeCli && (
        <AiModelSelect provider={provider} disabled={running} onChoose={choice.chooseModel} />
      )}
    </>
  )
}
