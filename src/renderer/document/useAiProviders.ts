import { useCallback, useEffect, useState } from 'react'
import type { IpcResponse, IpcResult } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { useWorkspace } from '../state/workspace.js'

export type AiProviderList = IpcResponse<typeof IpcChannel.AiProvidersList>
export type AiProviderInfo = AiProviderList['providers'][number]

export interface AiProviders {
  readonly list: AiProviderList | null
  readonly update: (call: Promise<IpcResult<AiProviderList>>) => Promise<boolean>
}

export function useAiProviders(configuring = false): AiProviders {
  const showError = useWorkspace((state) => state.showError)
  const [list, setList] = useState<AiProviderList | null>(null)

  const update = useCallback(
    async (call: Promise<IpcResult<AiProviderList>>): Promise<boolean> => {
      const result = await call
      if (!result.ok) {
        showError(result.error)
        return false
      }
      setList(result.data)
      return true
    },
    [showError],
  )

  useEffect(() => {
    if (!configuring) void update(window.api.ai.listProviders({}))
  }, [update, configuring])

  return { list, update }
}

export function authorOf(provider: AiProviderInfo, model: string | undefined): string {
  return model !== undefined && model !== '' ? model : provider.label
}
