import { randomUUID } from 'node:crypto'
import type { WebContents } from 'electron'
import type { AiProposal } from '@services/ai/proposals.js'
import type { IpcRequest, PushPayload } from '@shared/ipc.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { MCP_DOCUMENT_REFUSED, McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { sendPush } from '../window.js'
import { DocumentUnavailable, type DocumentWindow } from './mcp-external.js'

/** Reading the paragraphs or putting in one change takes the window well under this. */
export const MCP_WINDOW_TIMEOUT_MS = 5_000

type Reply = IpcRequest<typeof IpcChannel.McpDocumentReply>
type Ask = Omit<PushPayload<typeof IpcChannel.McpDocumentRequest>, 'requestId'>

const REFUSALS: Readonly<Record<McpDocumentRefusal, string>> = {
  [McpDocumentRefusal.NoTextDocument]:
    'No text document is open in Librevia (a spreadsheet, a PDF or nothing is): open one and try again.',
  [McpDocumentRefusal.ReadOnly]: 'The open document is read-only: it takes no revisions or comments.',
}
const NO_WINDOW = 'Librevia has no window open.'
const NO_ANSWER = 'The Librevia window did not answer in time; try again.'

function accepted(reply: Reply): Reply {
  if (reply.ask === MCP_DOCUMENT_REFUSED) throw new DocumentUnavailable(REFUSALS[reply.refusal])
  return reply
}

/** The main window, asked each time: the document may have changed since the last call. */
export class RendererDocument implements DocumentWindow {
  private contents: WebContents | null = null
  private readonly pending = new Map<string, (reply: Reply) => void>()

  attach(contents: WebContents): void {
    this.contents = contents
    contents.once('destroyed', () => {
      if (this.contents === contents) this.contents = null
    })
  }

  async snapshot(): Promise<readonly string[]> {
    const reply = await this.ask({ ask: McpDocumentAsk.Snapshot })
    return reply.ask === McpDocumentAsk.Snapshot ? reply.paragraphs : []
  }

  async apply(proposal: AiProposal, author: string): Promise<boolean> {
    const reply = await this.ask({ ask: McpDocumentAsk.Apply, proposal, author })
    return reply.ask === McpDocumentAsk.Apply && reply.applied
  }

  /** Only from the window that was asked, and only for a question still waiting. */
  receive(reply: Reply, sender: WebContents): boolean {
    const resolve = sender === this.contents ? this.pending.get(reply.requestId) : undefined
    resolve?.(reply)
    return resolve !== undefined
  }

  private ask(ask: Ask): Promise<Reply> {
    const contents = this.contents
    if (contents === null || contents.isDestroyed()) {
      return Promise.reject(new DocumentUnavailable(NO_WINDOW))
    }
    const requestId = randomUUID()
    return new Promise<Reply>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId)
        reject(new DocumentUnavailable(NO_ANSWER))
      }, MCP_WINDOW_TIMEOUT_MS)
      this.pending.set(requestId, (reply) => {
        clearTimeout(timer)
        this.pending.delete(requestId)
        resolve(reply)
      })
      sendPush(contents, IpcChannel.McpDocumentRequest, { requestId, ...ask })
    }).then(accepted)
  }
}
