import { useCallback, useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { Transaction } from '@tiptap/pm/state'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import { MAX_AI_SELECTION_LENGTH } from '@shared/limits.js'
import { AiFormat, type AiAction } from '@shared/types.js'
import type { Language } from '@shared/i18n/index.js'
import { useWorkspace } from '../state/workspace.js'
import {
  aiTarget,
  buildInsertBelow,
  buildTrackedReplacement,
  documentAround,
  rangeChanged,
  selectedMarkdown,
  type AiRange,
  type AiTargetProblem,
} from './ai-apply.js'

export type AiPhase = 'idle' | 'running' | 'done' | 'canceled' | 'declined' | 'failed'

export interface AiRunState {
  readonly phase: AiPhase
  readonly text: string
  readonly format: AiFormat
  readonly changed: boolean
  readonly model?: string | undefined
  readonly costUsd?: number | undefined
}

export interface AiStartOptions {
  readonly providerId: string
  readonly model: string
  readonly action: AiAction
  readonly targetLanguage?: Language | undefined
  readonly instruction?: string | undefined
  readonly format: AiFormat
}

interface ActiveRun {
  readonly requestId: string
  range: AiRange
  readonly snapshot: string
}

const IDLE: AiRunState = { phase: 'idle', text: '', format: AiFormat.Plain, changed: false }

export interface AiRequest {
  readonly state: AiRunState
  readonly start: (options: AiStartOptions) => Promise<AiTargetProblem | null>
  readonly cancel: () => void
  readonly apply: (author: string) => void
  readonly insertBelow: (author: string) => void
}

function useStreamedText(runRef: { readonly current: ActiveRun | null }, setState: SetState): void {
  useEffect(
    () =>
      window.api.ai.onStream((payload) => {
        const parsed = pushContracts[IpcChannel.AiStream].safeParse(payload)
        if (!parsed.success || parsed.data.requestId !== runRef.current?.requestId) return
        const { delta } = parsed.data
        setState((state) => ({ ...state, text: state.text + delta }))
      }),
    [runRef, setState],
  )
}

type SetState = (update: (state: AiRunState) => AiRunState) => void

function useTrackedRange(
  editor: Editor,
  runRef: { readonly current: ActiveRun | null },
  setState: SetState,
): void {
  useEffect(() => {
    const follow = ({ transaction }: { transaction: Transaction }): void => {
      const run = runRef.current
      if (run === null || !transaction.docChanged) return
      run.range = {
        from: transaction.mapping.map(run.range.from, 1),
        to: transaction.mapping.map(run.range.to, -1),
      }
      const changed = rangeChanged(editor.state.doc, run.range, run.snapshot)
      setState((state) => (state.changed === changed ? state : { ...state, changed }))
    }
    editor.on('transaction', follow)
    return () => {
      editor.off('transaction', follow)
    }
  }, [editor, runRef, setState])
}

export function useAiRequest(editor: Editor): AiRequest {
  const showError = useWorkspace((state) => state.showError)
  const [state, setState] = useState<AiRunState>(IDLE)
  const runRef = useRef<ActiveRun | null>(null)
  useStreamedText(runRef, setState)
  useTrackedRange(editor, runRef, setState)

  useEffect(
    () => () => {
      const run = runRef.current
      if (run !== null) void window.api.ai.cancel({ requestId: run.requestId })
    },
    [],
  )

  const start = useCallback(
    async (options: AiStartOptions): Promise<AiTargetProblem | null> => {
      const target = aiTarget(editor.state)
      if ('problem' in target) return target.problem
      const { doc } = editor.state
      const text = options.format === AiFormat.Markdown ? selectedMarkdown(doc, target.range) : target.text
      if (text.length > MAX_AI_SELECTION_LENGTH) return 'tooLong'
      const requestId = crypto.randomUUID()
      runRef.current = { requestId, range: target.range, snapshot: target.text }
      setState({ ...IDLE, phase: 'running', format: options.format })
      const around = documentAround(doc, target.range)
      const result = await window.api.ai.generate({ requestId, ...options, text, ...around })
      if (runRef.current?.requestId !== requestId) return null
      if (!result.ok) {
        showError(result.error)
        setState((current) => ({ ...current, phase: 'failed' }))
      } else {
        const { status, model, costUsd } = result.data
        setState((current) => ({ ...current, phase: status, model, costUsd }))
      }
      return null
    },
    [editor, showError],
  )

  const cancel = useCallback(() => {
    const run = runRef.current
    if (run !== null) void window.api.ai.cancel({ requestId: run.requestId })
  }, [])

  const dispatch = useCallback(
    (build: typeof buildTrackedReplacement, author: string) => {
      const run = runRef.current
      if (run === null) return
      const response = { text: state.text, format: state.format }
      const tr = build(editor.state, run.range, response, author, new Date())
      if (tr !== null) editor.view.dispatch(tr)
      editor.view.focus()
    },
    [editor, state.text, state.format],
  )

  return {
    state,
    start,
    cancel,
    apply: (author) => dispatch(buildTrackedReplacement, author),
    insertBelow: (author) => dispatch(buildInsertBelow, author),
  }
}
