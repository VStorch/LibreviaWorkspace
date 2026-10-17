import { useCallback, useRef, useState } from 'react'
import type { Range } from '@services/spreadsheet/edit.js'
import type { AiSheetTask } from '@shared/types.js'
import { currentPreferences } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import { useAiStream, type AiPhase } from '../document/useAiRequest.js'
import { useCancelOnUnmount } from '../ai/useCancelOnUnmount.js'

export interface SheetAiRun {
  readonly phase: AiPhase
  readonly task: AiSheetTask
  readonly text: string
  /** The cells the answer is for, on the sheet that was active when asking. */
  readonly target: Range
  readonly sheetIndex: number
  readonly model?: string | undefined
  readonly costUsd?: number | undefined
}

export interface SheetAiAsk {
  readonly providerId: string
  readonly model: string
  readonly task: AiSheetTask
  readonly instruction: string
  readonly formula?: string | undefined
  readonly context: string
  readonly rowCount?: number | undefined
  readonly target: Range
  readonly sheetIndex: number
}

export interface SheetAi {
  readonly run: SheetAiRun | null
  readonly ask: (request: SheetAiAsk) => Promise<void>
  readonly cancel: () => void
}

export function useSheetAi(): SheetAi {
  const showError = useWorkspace((state) => state.showError)
  const [run, setRun] = useState<SheetAiRun | null>(null)
  const requestId = useRef<string | null>(null)
  const activeRequestId = useCallback(() => requestId.current ?? undefined, [])
  const append = useCallback(
    (delta: string) =>
      setRun((current) => (current === null ? current : { ...current, text: current.text + delta })),
    [],
  )
  useAiStream(activeRequestId, append)
  useCancelOnUnmount(requestId)

  const ask = useCallback(
    async ({ target, sheetIndex, ...request }: SheetAiAsk): Promise<void> => {
      const id = crypto.randomUUID()
      requestId.current = id
      setRun({ phase: 'running', task: request.task, text: '', target, sheetIndex })
      const language = currentPreferences().language
      const result = await window.api.ai.sheet({ requestId: id, language, ...request })
      if (requestId.current !== id) return
      if (!result.ok) {
        showError(result.error)
        setRun((current) => current && { ...current, phase: 'failed' })
        return
      }
      const { status, model, costUsd } = result.data
      setRun((current) => current && { ...current, phase: status, model, costUsd })
    },
    [showError],
  )

  const cancel = useCallback(() => {
    if (requestId.current !== null) void window.api.ai.cancel({ requestId: requestId.current })
  }, [])

  return { run, ask, cancel }
}
