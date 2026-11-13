import type { IpcRequest } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { OpenMode } from '@shared/types.js'
import { currentPreferences } from './preferences.js'

/** One file per window, as in Word: a file opened from a window that shows one gets its own. */
export function opensWindows(): boolean {
  return currentPreferences().openMode === OpenMode.Windows
}

/** Whether main took it: another window shows the file, or a new one opens for it. */
export async function openInAnotherWindow(
  request: IpcRequest<typeof IpcChannel.WindowOpen>,
): Promise<boolean> {
  const result = await window.api.window.open(request)
  return result.ok && result.data.placed
}
