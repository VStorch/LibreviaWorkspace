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

  it('sends the window the deadline and waits a bit beyond it, so a change already made is not reported lost', async () => {
    const { document, contents } = attached()
    const applying = document.apply(PROPOSAL, 'Claude Code (MCP)')
    const [request] = sent
    expect(request?.deadline).toBe(Date.now() + MCP_WINDOW_TIMEOUT_MS)

    vi.advanceTimersByTime(MCP_WINDOW_TIMEOUT_MS + MCP_REPLY_GRACE_MS / 2)
    document.receive({ requestId: request!.requestId, ask: McpDocumentAsk.Apply, applied: true }, contents)

    await expect(applying).resolves.toBe(true)
  })

  it('the window refuses the expired request and the client learns nothing changed', async () => {
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

  it('with no response after the deadline and slack, the request fails and the late response is ignored', async () => {
    const { document, contents } = attached()
    const applying = document.apply(PROPOSAL, 'Claude Code (MCP)')
    const failure = expect(applying).rejects.toBeInstanceOf(DocumentUnavailable)
    vi.advanceTimersByTime(MCP_WINDOW_TIMEOUT_MS + MCP_REPLY_GRACE_MS)
    await failure

    const late = { requestId: sent[0]!.requestId, ask: McpDocumentAsk.Apply, applied: true } as const
    expect(document.receive(late, contents)).toBe(false)
  })

  it('the sheet tool goes to the window with the request and the author, and the response returns to the client', async () => {
    const { document, contents } = attached()
    const call = { id: '1', name: 'read_cells', arguments: '{"range":"A1"}' }
    const running = document.runSheetTool(call, 'Editor X (MCP)')
    expect(sent[0]).toMatchObject({
      ask: McpDocumentAsk.SheetTool,
      call: { name: call.name, arguments: call.arguments },
      author: 'Editor X (MCP)',
    })
    const reply = {
      requestId: sent[0]!.requestId,
      ask: McpDocumentAsk.SheetTool,
      content: 'A1: 1',
      isError: false,
    } as const
    document.receive(reply, contents)
    await expect(running).resolves.toEqual({ content: 'A1: 1', isError: false })
  })

  it('without an open sheet, the client is told', async () => {
    const { document, contents } = attached()
    const running = document.runSheetTool({ id: '1', name: 'list_sheets', arguments: '{}' }, 'X (MCP)')
    const reply = {
      requestId: sent[0]!.requestId,
      ask: MCP_DOCUMENT_REFUSED,
      refusal: McpDocumentRefusal.NoSpreadsheet,
    } as const
    document.receive(reply, contents)
    await expect(running).rejects.toThrow(/No spreadsheet is open/u)
  })
})
