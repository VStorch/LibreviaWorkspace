import { useCallback, useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { AiProposal } from '@services/ai/proposals.js'
import { AiRole, type AiTurn } from '@services/ai/provider.js'
import { MAX_AI_CHAT_MESSAGES } from '@shared/limits.js'
import { AiFormat } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'
import { buildInsertBelow, chatDocument } from './ai-apply.js'
import { chatParagraphs } from './ai-proposals.js'
import { applyProposals } from './proposal-commands.js'
import { useAiStream, type AiPhase } from './useAiRequest.js'

export interface AiChatMessage {
  readonly id: string
  readonly role: AiRole
  readonly text: string
  /** Always `done` for what the person wrote. */
  readonly phase: AiPhase
  readonly model?: string | undefined
  readonly costUsd?: number | undefined
  readonly proposals?: readonly ChatProposal[] | undefined
  /** Why a failed reply has no answer, as main worded it. */
  readonly error?: string | undefined
}

export type ProposalState = 'open' | 'applied' | 'discarded' | 'changed'

export type ChatProposal = AiProposal & { readonly state: ProposalState }

export interface AiChatTarget {
  readonly providerId: string
  readonly model: string
  /** The model may propose edits and comments: the document goes numbered by paragraph. */
  readonly proposals: boolean
}

/** In memory only, for the open document: nothing of it is saved anywhere. */
export interface AiChat {
  readonly messages: readonly AiChatMessage[]
  readonly running: boolean
  readonly send: (text: string, target: AiChatTarget) => Promise<void>
  readonly cancel: () => void
  readonly clear: () => void
  readonly insertBelow: (message: AiChatMessage, author: string) => void
  /** Only the open ones among `ids`; the rest keep their state. */
  readonly apply: (messageId: string, ids: readonly string[], author: string) => void
  readonly discard: (messageId: string, id: string) => void
}

type SetMessages = (update: (messages: readonly AiChatMessage[]) => readonly AiChatMessage[]) => void

function historyOf(messages: readonly AiChatMessage[]): AiTurn[] {
  return messages
    .filter((message) => message.text !== '')
    .map(({ role, text }) => ({ role, text }))
    .slice(-MAX_AI_CHAT_MESSAGES)
}

function updated(
  messages: readonly AiChatMessage[],
  id: string,
  change: (message: AiChatMessage) => Partial<AiChatMessage>,
): readonly AiChatMessage[] {
  return messages.map((message) => (message.id === id ? { ...message, ...change(message) } : message))
}

function withStates(
  message: AiChatMessage,
  states: ReadonlyMap<string, ProposalState>,
): Partial<AiChatMessage> {
  return {
    proposals: message.proposals?.map((proposal) => ({
      ...proposal,
      state: states.get(proposal.id) ?? proposal.state,
    })),
  }
}

function useProposals(
  editor: Editor | null,
  messages: readonly AiChatMessage[],
  setMessages: SetMessages,
): Pick<AiChat, 'apply' | 'discard'> {
  const apply = useCallback(
    (messageId: string, ids: readonly string[], author: string) => {
      const open = messages
        .find((message) => message.id === messageId)
        ?.proposals?.filter((proposal) => proposal.state === 'open' && ids.includes(proposal.id))
      if (editor === null || open === undefined || open.length === 0) return
      const outcome = applyProposals(editor, open, author)
      const states = new Map<string, ProposalState>([
        ...outcome.applied.map((id) => [id, 'applied'] as const),
        ...outcome.changed.map((id) => [id, 'changed'] as const),
      ])
      setMessages((current) => updated(current, messageId, (message) => withStates(message, states)))
    },
    [editor, messages, setMessages],
  )
  const discard = useCallback(
    (messageId: string, id: string) => {
      const states = new Map<string, ProposalState>([[id, 'discarded']])
      setMessages((current) => updated(current, messageId, (message) => withStates(message, states)))
    },
    [setMessages],
  )
  return { apply, discard }
}

function useReplyStream(activeRef: { readonly current: string | null }, setMessages: SetMessages): void {
  const activeRequestId = useCallback(() => activeRef.current ?? undefined, [activeRef])
  const append = useCallback(
    (delta: string) => {
      const id = activeRef.current
      if (id !== null)
        setMessages((messages) => updated(messages, id, (reply) => ({ text: reply.text + delta })))
    },
    [activeRef, setMessages],
  )
  useAiStream(activeRequestId, append)
}

function useCancelOnUnmount(activeRef: { readonly current: string | null }): void {
  useEffect(
    () => () => {
      const requestId = activeRef.current
      if (requestId !== null) void window.api.ai.cancel({ requestId })
    },
    [activeRef],
  )
}

export function useAiChat(editor: Editor | null): AiChat {
  const showError = useWorkspace((state) => state.showError)
  const [messages, setMessages] = useState<readonly AiChatMessage[]>([])
  const activeRef = useRef<string | null>(null)
  useReplyStream(activeRef, setMessages)
  useCancelOnUnmount(activeRef)

  const send = useCallback(
    async (text: string, target: AiChatTarget): Promise<void> => {
      const question = text.trim()
      if (editor === null || activeRef.current !== null || question === '') return
      const requestId = crypto.randomUUID()
      const asked: AiChatMessage = {
        id: crypto.randomUUID(),
        role: AiRole.User,
        text: question,
        phase: 'done',
      }
      const reply: AiChatMessage = { id: requestId, role: AiRole.Assistant, text: '', phase: 'running' }
      activeRef.current = requestId
      setMessages((current) => [...current, asked, reply].slice(-MAX_AI_CHAT_MESSAGES))
      const history = historyOf([...messages, asked])
      const result = await window.api.ai.chat({
        requestId,
        providerId: target.providerId,
        model: target.model,
        messages: history,
        ...chatDocument(editor.state),
        proposals: target.proposals,
        paragraphs: target.proposals ? chatParagraphs(editor.state.doc) : [],
      })
      if (activeRef.current !== requestId) return
      activeRef.current = null
      if (!result.ok) showError(result.error)
      const {
        status: phase,
        model,
        costUsd,
        proposals,
      } = result.ok ? result.data : { status: 'failed' as const }
      const opened = proposals?.map((proposal): ChatProposal => ({ ...proposal, state: 'open' }))
      const error = result.ok ? undefined : result.error.message
      setMessages((current) =>
        updated(current, requestId, () => ({ phase, model, costUsd, proposals: opened, error })),
      )
    },
    [editor, messages, showError],
  )

  const cancel = useCallback(() => {
    const requestId = activeRef.current
    if (requestId !== null) void window.api.ai.cancel({ requestId })
  }, [])

  const clear = useCallback(() => {
    cancel()
    activeRef.current = null
    setMessages([])
  }, [cancel])

  const insertBelow = useCallback(
    (message: AiChatMessage, author: string) => {
      if (editor === null) return
      const { state } = editor
      const range = { from: state.selection.from, to: state.selection.to }
      const response = { text: message.text, format: AiFormat.Markdown }
      const tr = buildInsertBelow(state, range, response, author, new Date())
      if (tr !== null) editor.view.dispatch(tr)
      editor.view.focus()
    },
    [editor],
  )

  const { apply, discard } = useProposals(editor, messages, setMessages)
  const running = messages.some((message) => message.phase === 'running')
  return { messages, running, send, cancel, clear, insertBelow, apply, discard }
}
