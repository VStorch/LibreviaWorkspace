import { useState } from 'react'
import type { IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import type { AiProviderInfo } from './useAiProviders.js'

type Model = IpcResponse<typeof IpcChannel.AiModelsList>['models'][number]

const AUTOMATIC = ''

/** The provider's list only comes on request: listing may reach the network and ask for consent. */
export function AiModelSelect({
  provider,
  disabled,
  onChoose,
}: {
  readonly provider: AiProviderInfo
  readonly disabled: boolean
  readonly onChoose: (model: string) => void
}): React.JSX.Element {
  const t = useT()
  const showError = useWorkspace((state) => state.showError)
  const [loaded, setLoaded] = useState<{ readonly providerId: string; readonly models: readonly Model[] }>()
  const loadModels = async (): Promise<void> => {
    const result = await window.api.ai.listModels({ providerId: provider.id })
    if (result.ok) setLoaded({ providerId: provider.id, models: result.data.models })
    else showError(result.error)
  }
  const chosen = provider.defaultModel ?? AUTOMATIC
  const listed = loaded?.providerId === provider.id ? loaded.models : []
  const options =
    chosen === AUTOMATIC || listed.some((model) => model.id === chosen)
      ? listed
      : [{ id: chosen, label: chosen }, ...listed]
  return (
    <div className="popover__row">
      <label className="popover__field">
        <span>{t('ai.panel.model')}</span>
        <select value={chosen} disabled={disabled} onChange={(event) => onChoose(event.target.value)}>
          <option value={AUTOMATIC}>{t('ai.providers.modelAuto')}</option>
          {options.map((model) => (
            <option key={model.id} value={model.id}>
              {model.label}
            </option>
          ))}
        </select>
      </label>
      <button type="button" className="btn" disabled={disabled} onClick={() => void loadModels()}>
        {t('ai.providers.loadModels')}
      </button>
    </div>
  )
}
