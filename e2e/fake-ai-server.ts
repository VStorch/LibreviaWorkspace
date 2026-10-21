import { createServer, type Server, type ServerResponse } from 'node:http'
import type { AddressInfo } from 'node:net'
import { expect } from '@playwright/test'
import { menu, type Session } from './app.js'

export const TEMPLATE = 'modelo-falso'
export const REPLY = ['Texto ', 'reescrito ', 'pelo ', 'modelo.']

const SLOW_CHUNKS = 20

/**
 * A fake OpenAI-compatible server on 127.0.0.1, answering slowly so the test sees the text arrive.
 * `lento` keeps the stream open long enough to cancel or edit during it; `derrubar` drops the
 * connection after the first piece.
 */
export interface FakeServer {
  readonly server: Server
  readonly baseUrl: string
  readonly requests: string[]
  readonly bodies: string[]
  reply: string[]
  /** With `tools` in the request and no tool result yet, the reply is these calls, in pieces. */
  calls: { name: string; arguments: unknown }[] | null
  closedEarly: number
  slow: boolean
  dropConnection: boolean
  /** Answers a request with `tools` as Ollama does for a model without them. */
  withoutTools: boolean
}

export async function startFakeServer(): Promise<FakeServer> {
  const fake = {
    requests: [] as string[],
    bodies: [] as string[],
    reply: REPLY,
    calls: null,
    closedEarly: 0,
    slow: false,
    dropConnection: false,
    withoutTools: false,
  } as FakeServer
  const server = createServer((req, res) => {
    fake.requests.push(`${req.method} ${req.url}`)
    let body = ''
    req.on('data', (chunk: Buffer) => (body += chunk.toString()))
    req.on('end', () => {
      if (body !== '') fake.bodies.push(body)
      if (req.url === '/v1/models') {
        res.writeHead(200, { 'content-type': 'application/json' })
        res.end(JSON.stringify({ data: [{ id: TEMPLATE }] }))
        return
      }
      reply(res, body)
    })
  })
  function chunksOf(body: string): string[] {
    const request = JSON.parse(body) as { tools?: unknown[]; messages: { role: string }[] }
    const firstRound = request.tools !== undefined && !request.messages.some((m) => m.role === 'tool')
    if (firstRound && fake.calls !== null) {
      return fake.calls.flatMap((call, index) => {
        const args = JSON.stringify(call.arguments)
        const half = Math.floor(args.length / 2)
        return [
          {
            index,
            id: `chamada-${index}`,
            type: 'function',
            function: { name: call.name, arguments: args.slice(0, half) },
          },
          { index, function: { arguments: args.slice(half) } },
        ].map((delta) => JSON.stringify({ choices: [{ delta: { tool_calls: [delta] } }] }))
      })
    }
    const texts = fake.slow ? Array.from({ length: SLOW_CHUNKS }, (_, i) => `${i} `) : fake.reply
    return texts.map((text) => JSON.stringify({ choices: [{ delta: { content: text } }] }))
  }
  function reply(res: ServerResponse, body: string): void {
    if (fake.withoutTools && (JSON.parse(body) as { tools?: unknown }).tools !== undefined) {
      res.writeHead(400, { 'content-type': 'application/json' })
      res.end(JSON.stringify({ error: { message: `${TEMPLATE} does not support tools` } }))
      return
    }
    res.writeHead(200, { 'content-type': 'text/event-stream' })
    const chunks = chunksOf(body)
    let i = 0
    const timer = setInterval(() => {
      if (fake.dropConnection && i === 1) {
        res.socket?.destroy()
        return
      }
      if (i < chunks.length) {
        res.write(`data: ${chunks[i++]}\n\n`)
        return
      }
      clearInterval(timer)
      res.end('data: [DONE]\n\n')
    }, 250)
    res.on('close', () => {
      clearInterval(timer)
      if (!res.writableEnded) fake.closedEarly++
    })
  }
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  const { port } = server.address() as AddressInfo
  return Object.assign(fake, { server, baseUrl: `http://127.0.0.1:${port}/v1` })
}

/** Through the setup dialog, as a person would: a server on 127.0.0.1 needs no consent. */
export async function configureAssistant(session: Session, baseUrl: string): Promise<void> {
  await menu(session, 'ai-providers')
  const dialog = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
  await dialog.getByLabel('Nome').fill('Servidor falso')
  await dialog.getByLabel('Endereço').fill(baseUrl)
  await dialog.getByRole('button', { name: 'Salvar' }).click()
  await expect(dialog.getByText('Servidor falso')).toBeVisible()
  await expect(dialog.getByText('Neste computador · Padrão')).toBeVisible()
  await dialog.getByRole('button', { name: 'Fechar' }).click()
}
