import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { AiProposal } from '@services/ai/proposals.js'
import { AiRole, type AiTurn } from '@services/ai/provider.js'
import { MAX_AI_CHAT_MESSAGES } from '@shared/limits.js'
import { AiFormat } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'
import { buildInsertBelow, chatDocument } from './ai-apply.js'
import { chatParagraphs } from './ai-proposals.js'
import { showAiDiff, type AiDiffEntry } from './extensions/ai-diff.js'
import { applyProposals } from './proposal-commands.js'
import { useAiStream, type AiPhase } from './useAiRequest.js'
import { useCancelOnUnmount } from '../ai/useCancelOnUnmount.js'

export interface AiChatMessage {
  readonly id: string
  readonly role: AiRole
  readonly text: string
  /** Always `done` for what the person wrote. */
  readonly phase: AiPhase
  readonly model?: string | undefined
  readonly costUsd?: number | undefined
  readonly proposals?: readonly ChatProposal[] | undefined
  /** The author of what a reply puts in the document when the model did not name itself. */
  readonly providerLabel?: string | undefined
  /** Why a failed reply has no answer, as main worded it. */
  readonly error?: string | undefined
}

export type ProposalState = 'open' | 'applied' | 'discarded' | 'changed'

export type ChatProposal = AiProposal & { readonly state: ProposalState }

export interface AiChatTarget {
  readonly providerId: string
  readonly model: string
  readonly providerLabel: string
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
  /**
   * Only the open ones among `ids`; the rest keep their state. With revision tracking off, they go
   * straight into the text; on, as a revision by the model.
   */
  readonly apply: (messageId: string, ids: readonly string[]) => void
  readonly discard: (messageId: string, ids: readonly string[]) => void
  /** Every open proposal of the conversation, to draw on the sheet. */
  readonly openProposals: readonly AiDiffEntry[]
  /** For a proposal drawn on the sheet, by its entry key. */
  readonly decide: (key: string, accept: boolean) => void
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

const authorOfReply = (message: AiChatMessage): string =>
  message.model !== undefined && message.model !== '' ? message.model : (message.providerLabel ?? '')

const KEY_SEPARATOR = ':'

const entryKey = (messageId: string, proposalId: string): string =>
  `${messageId}${KEY_SEPARATOR}${proposalId}`

function openEntries(messages: readonly AiChatMessage[]): AiDiffEntry[] {
  return messages.flatMap((message) =>
    (message.proposals ?? [])
      .filter((proposal) => proposal.state === 'open')
      .map(({ state: _state, ...proposal }) => ({ key: entryKey(message.id, proposal.id), proposal })),
  )
}

function useProposals(
  editor: Editor | null,
  messages: readonly AiChatMessage[],
  setMessages: SetMessages,
): Pick<AiChat, 'apply' | 'discard' | 'openProposals' | 'decide'> {
  const apply = useCallback(
    (messageId: string, ids: readonly string[]) => {
      const message = messages.find((candidate) => candidate.id === messageId)
      const open = message?.proposals?.filter(
        (proposal) => proposal.state === 'open' && ids.includes(proposal.id),
      )
      if (editor === null || message === undefined || open === undefined || open.length === 0) return
      const direct = useWorkspace.getState().trackChanges !== true
      const outcome = applyProposals(editor, open, authorOfReply(message), direct)
      const states = new Map<string, ProposalState>([
        ...outcome.applied.map((id) => [id, 'applied'] as const),
        ...outcome.changed.map((id) => [id, 'changed'] as const),
      ])
      setMessages((current) => updated(current, messageId, (message) => withStates(message, states)))
    },
    [editor, messages, setMessages],
  )
  const discard = useCallback(
    (messageId: string, ids: readonly string[]) => {
      const states = new Map<string, ProposalState>(ids.map((id) => [id, 'discarded']))
      setMessages((current) => updated(current, messageId, (message) => withStates(message, states)))
    },
    [setMessages],
  )
  const decide = useCallback(
    (key: string, accept: boolean) => {
      const at = key.lastIndexOf(KEY_SEPARATOR)
      const [messageId, id] = [key.slice(0, at), key.slice(at + 1)]
      if (accept) apply(messageId, [id])
      else discard(messageId, [id])
    },
    [apply, discard],
  )
  const openProposals = useMemo(() => openEntries(messages), [messages])
  return { apply, discard, openProposals, decide }
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
      const reply: AiChatMessage = {
        id: requestId,
        role: AiRole.Assistant,
        text: '',
        phase: 'running',
        providerLabel: target.providerLabel,
      }
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

  const proposals = useProposals(editor, messages, setMessages)
  useAiDiff(editor, proposals)
  const running = messages.some((message) => message.phase === 'running')
  return { messages, running, send, cancel, clear, insertBelow, ...proposals }
}

/** The open proposals, drawn on the sheet until each is accepted or rejected. */
function useAiDiff(
  editor: Editor | null,
  { openProposals, decide }: Pick<AiChat, 'openProposals' | 'decide'>,
): void {
  const readOnly = useWorkspace((state) => state.readOnly)
  useEffect(() => {
    if (editor === null || editor.isDestroyed) return
    showAiDiff(editor.view, { entries: openProposals, readOnly, decide })
  }, [editor, openProposals, readOnly, decide])
}
