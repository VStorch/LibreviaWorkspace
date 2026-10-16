import { AppError, ErrorCode } from '@shared/errors.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { t } from '../i18n.js'

const HTTP_UNAUTHORIZED = 401
const HTTP_FORBIDDEN = 403
const HTTP_TOO_MANY_REQUESTS = 429
const MAX_DETAIL_LENGTH = 300

export function aiError(code: ErrorCode, key: MessageKey, detail?: string): AppError {
  return new AppError(
    code,
    t(key),
    detail === undefined || detail === '' ? undefined : detail.slice(0, MAX_DETAIL_LENGTH),
  )
}

export function errorForStatus(status: number): AppError {
  if (status === HTTP_UNAUTHORIZED || status === HTTP_FORBIDDEN) {
    return aiError(ErrorCode.AiAuthFailed, 'errors.ai.auth')
  }
  if (status === HTTP_TOO_MANY_REQUESTS) return aiError(ErrorCode.AiRateLimited, 'errors.ai.rateLimited')
  return aiError(ErrorCode.AiFailed, 'errors.ai.failed', `HTTP ${status}`)
}

export const refusedError = (): AppError => aiError(ErrorCode.AiRefused, 'errors.ai.refused')
export const unreachableError = (): AppError => aiError(ErrorCode.AiUnreachable, 'errors.ai.unreachable')
export const toolCallsTooLargeError = (): AppError =>
  aiError(ErrorCode.AiFailed, 'errors.ai.toolCallsTooLarge')
