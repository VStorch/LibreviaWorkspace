import { IpcChannel } from '@shared/ipc-channels.js'
import { AiRequests } from '../ai/requests.js'
import { ensureConsent, generate, providerFor } from '../ai/assistant.js'
import { discoverProviders } from '../ai/discovery.js'
import { AiModelCache } from '../ai/model-cache.js'
import { aiSettings } from '../ai/settings.js'
import { confirmAiConsent } from '../dialogs.js'
import { sendPush } from '../window.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

const MODELS_TIMEOUT_MS = 15_000

/** The only handlers that reach the network, and only after the person sets up a provider. */
export function registerAiHandlers(): void {
  const requests = new AiRequests()
  const models = new AiModelCache()

  handle(IpcChannel.AiProvidersList, () => aiSettings().list())
  handle(IpcChannel.AiProviderSave, (payload) => {
    models.forget(payload.id)
    return aiSettings().save(payload)
  })
  handle(IpcChannel.AiProviderRemove, (payload) => {
    models.forget(payload.providerId)
    return aiSettings().remove(payload.providerId)
  })
  handle(IpcChannel.AiDiscover, async () => ({ candidates: await discoverProviders() }))

  handle(IpcChannel.AiModelsList, async (payload, event) => {
    const settings = aiSettings()
    const stored = settings.find(payload.providerId)
    const confirm = (): Promise<boolean> => confirmAiConsent(windowOf(event), stored.label, stored.baseUrl)
    if (!(await ensureConsent(settings, stored, confirm))) return { models: [] }
    const provider = await providerFor(settings, stored)
    const listed = await provider.listModels(AbortSignal.timeout(MODELS_TIMEOUT_MS))
    return { models: [...models.remember(stored.id, listed)] }
  })

  handle(IpcChannel.AiGenerate, async (payload, event) => {
    const signal = requests.begin(payload.requestId, event.sender)
    try {
      return await generate(payload, {
        settings: aiSettings(),
        signal,
        push: (delta) => {
          if (!event.sender.isDestroyed())
            sendPush(event.sender, IpcChannel.AiStream, { requestId: payload.requestId, delta })
        },
        models,
        confirmConsent: (provider) => confirmAiConsent(windowOf(event), provider.label, provider.baseUrl),
      })
    } finally {
      requests.end(payload.requestId)
    }
  })

  handle(IpcChannel.AiCancel, (payload, event) => ({
    canceled: requests.cancel(payload.requestId, event.sender.id),
  }))
}
