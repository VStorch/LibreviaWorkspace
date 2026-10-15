import type { IpcMainInvokeEvent } from 'electron'
import { AiProviderKind, type AiModel } from '@services/ai/provider.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { AiRequests } from '../ai/requests.js'
import { ensureConsent, generate, providerFor } from '../ai/assistant.js'
import { discoverProviders } from '../ai/discovery.js'
import { AiModelCache } from '../ai/model-cache.js'
import { aiSettings } from '../ai/settings.js'
import { confirmAiConsent, showAgentExecutableDialog } from '../dialogs.js'
import { sendPush } from '../window.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

const MODELS_TIMEOUT_MS = 15_000

/** Ends with the window that asked: an agent started only to list its models dies with it too. */
async function listModels(
  providerId: string,
  event: IpcMainInvokeEvent,
  models: AiModelCache,
): Promise<{ models: AiModel[] }> {
  const settings = aiSettings()
  const stored = settings.find(providerId)
  const confirm = (): Promise<boolean> => confirmAiConsent(windowOf(event), stored.label, stored.baseUrl)
  if (!(await ensureConsent(settings, stored, confirm))) return { models: [] }
  const provider = await providerFor(settings, stored)
  const closed = new AbortController()
  const abort = (): void => closed.abort()
  event.sender.once('destroyed', abort)
  const load = (): Promise<AiModel[]> =>
    provider.listModels(AbortSignal.any([AbortSignal.timeout(MODELS_TIMEOUT_MS), closed.signal]))
  try {
    // An agent is a whole process started just to read its list: once per provider is enough.
    if (stored.kind === AiProviderKind.Acp) return { models: [...(await models.get(stored.id, load))] }
    return { models: [...models.remember(stored.id, await load())] }
  } finally {
    event.sender.off('destroyed', abort)
  }
}

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
  handle(IpcChannel.AiModelChoose, (payload) => aiSettings().chooseModel(payload.providerId, payload.model))
  handle(IpcChannel.AiPreferencesSave, (payload) => aiSettings().savePreferences(payload))
  handle(IpcChannel.AiAgentChooseExecutable, async (payload, event) => {
    const command = await showAgentExecutableDialog(windowOf(event))
    if (command === null) return { ...aiSettings().list(), chosen: false }
    models.forget(payload.id)
    return { ...aiSettings().saveAgentExecutable(payload, command), chosen: true }
  })
  handle(IpcChannel.AiDiscover, async () => ({ candidates: await discoverProviders() }))

  handle(IpcChannel.AiModelsList, (payload, event) => listModels(payload.providerId, event, models))

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
