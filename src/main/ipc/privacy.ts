import { IpcChannel } from '@shared/ipc-channels.js'
import { assertPathAuthorized } from '../fs/paths.js'
import { inspectPersonalData, removePersonalData } from '../privacy.js'
import { sidecar } from '../sidecar/index.js'
import { handle } from './registry.js'

/** Only a file the person opened or saved: a Word template's path is never authorized. */
export function registerPrivacyHandlers(): void {
  handle(IpcChannel.PrivacyInspect, (payload) =>
    inspectPersonalData(sidecar(), assertPathAuthorized(payload.path)),
  )

  handle(IpcChannel.PrivacyRemove, (payload) =>
    removePersonalData(sidecar(), assertPathAuthorized(payload.path), payload.categories),
  )
}
