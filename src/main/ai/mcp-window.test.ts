import type { WebContents } from 'electron'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AiProposalKind, type AiProposal } from '@services/ai/proposals.js'
import type { PushPayload } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { MCP_DOCUMENT_REFUSED, McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { DocumentUnavailable } from './mcp-external.js'
import { MCP_REPLY_GRACE_MS, MCP_WINDOW_TIMEOUT_MS, RendererDocument } from './mcp-window.js'

type Sent = PushPayload<typeof IpcChannel.McpDocumentRequest>

const sent = vi.hoisted(() => [] as Sent[])

vi.mock('../window.js', () => ({
  sendPush: (_contents: unknown, _channel: unknown, payload: Sent) => sent.push(payload),
}))

const PROPOSAL: AiProposal = {
  id: 'p1',
  kind: AiProposalKind.Edit,
  paragraph: 1,
  paragraphText: 'os menino foi',
  find: 'menino foi',
  replace: 'meninos foram',
}

function attached(): { document: RendererDocument; contents: WebContents } {
  const contents = { once: () => undefined, isDestroyed: () => false } as unknown as WebContents
  const document = new RendererDocument()
  document.attach(contents)
  return { document, contents }
}

describe('RendererDocument', () => {
  beforeEach(() => {
    vi.useFakeTimers()
    sent.length = 0
  })
  afterEach(() => vi.useRealTimers())

  it('manda à janela o prazo e espera um pouco além dele, para não dar como perdida uma mudança feita', async () => {
    const { document, contents } = attached()
    const applying = document.apply(PROPOSAL, 'Claude Code (MCP)')
    const [request] = sent
    expect(request?.deadline).toBe(Date.now() + MCP_WINDOW_TIMEOUT_MS)

    vi.advanceTimersByTime(MCP_WINDOW_TIMEOUT_MS + MCP_REPLY_GRACE_MS / 2)
    document.receive({ requestId: request!.requestId, ask: McpDocumentAsk.Apply, applied: true }, contents)

    await expect(applying).resolves.toBe(true)
  })

  it('a janela recusa o pedido vencido e o cliente fica sabendo que nada mudou', async () => {
    const { document, contents } = attached()
    const applying = document.apply(PROPOSAL, 'Claude Code (MCP)')
    const reply = {
      requestId: sent[0]!.requestId,
      ask: MCP_DOCUMENT_REFUSED,
      refusal: McpDocumentRefusal.Expired,
    } as const
    document.receive(reply, contents)

    await expect(applying).rejects.toThrow(/nothing was changed/u)
    await expect(applying).rejects.toBeInstanceOf(DocumentUnavailable)
  })

  it('sem resposta depois do prazo e da folga, o pedido falha e a resposta atrasada é ignorada', async () => {
    const { document, contents } = attached()
    const applying = document.apply(PROPOSAL, 'Claude Code (MCP)')
    const failure = expect(applying).rejects.toBeInstanceOf(DocumentUnavailable)
    vi.advanceTimersByTime(MCP_WINDOW_TIMEOUT_MS + MCP_REPLY_GRACE_MS)
    await failure

    const late = { requestId: sent[0]!.requestId, ask: McpDocumentAsk.Apply, applied: true } as const
    expect(document.receive(late, contents)).toBe(false)
  })
})
