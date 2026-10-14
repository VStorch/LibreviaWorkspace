import { takeLines, type SplitLines } from '@services/ai/lines.js'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_STREAM_LINE_LENGTH } from '@shared/limits.js'
import { aiError } from './errors.js'

export function takeBoundedLines(buffer: string): SplitLines {
  const split = takeLines(buffer)
  if (split.rest.length > MAX_AI_STREAM_LINE_LENGTH)
    throw aiError(ErrorCode.AiFailed, 'errors.ai.lineTooLong')
  return split
}
