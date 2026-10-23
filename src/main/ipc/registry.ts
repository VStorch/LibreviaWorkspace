import { ipcMain, type IpcMainInvokeEvent } from 'electron'
import type { z } from 'zod'
import type { InvocableIpcChannel } from '@shared/ipc-channels.js'
import { ipcContracts, type IpcRequest, type IpcResponse, type IpcResult } from '@shared/ipc.js'
import { AppError, ErrorCode, toSerializedError } from '@shared/errors.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

function fieldDetail(error: z.ZodError, key: MessageKey): string | undefined {
  const field = error.issues[0]?.path.join('.')
  return field === undefined || field === '' ? undefined : t(key, { field })
}

/**
 * The request is validated against the channel schema: a renderer compromised by a document cannot
 * ask for an operation the contract does not allow. The response too, because main reads disk and
 * program output, and the schema is where the limits are.
 */
export function handle<C extends InvocableIpcChannel>(
  channel: C,
  handler: (payload: IpcRequest<C>, event: IpcMainInvokeEvent) => Promise<IpcResponse<C>> | IpcResponse<C>,
): void {
  ipcMain.handle(channel, async (event, rawPayload: unknown): Promise<IpcResult<IpcResponse<C>>> => {
    try {
      const contract = ipcContracts[channel]
      const parsed = contract.request.safeParse(rawPayload)

      if (!parsed.success) {
        throw new AppError(
          ErrorCode.InvalidRequest,
          t('errors.ipc.invalidData'),
          fieldDetail(parsed.error, 'errors.ipc.requestField'),
        )
      }

      const data = await handler(parsed.data as IpcRequest<C>, event)

      const replied = contract.response.safeParse(data)
      if (!replied.success) {
        // Our defect, not the user's: `Internal`, with the detail in the log.
        throw new AppError(
          ErrorCode.Internal,
          t('errors.ipc.unrecognizedEnd'),
          fieldDetail(replied.error, 'errors.ipc.responseField'),
        )
      }

      // The validated value, not the raw one: the schema applies defaults and drops undeclared
      // fields.
      return { ok: true, data: replied.data as IpcResponse<C> }
    } catch (cause) {
      console.error(`[ipc] channel ${channel} failed:`, cause)
      return { ok: false, error: toSerializedError(cause, editorPreferences().language) }
    }
  })
}
