import { createServer, type Server, type ServerResponse } from 'node:http'
import type { AddressInfo } from 'node:net'
import { expect } from '@playwright/test'
import { menu, type Session } from './app.js'

export const MODELO = 'modelo-falso'
export const RESPOSTA = ['Texto ', 'reescrito ', 'pelo ', 'modelo.']

const PEDACOS_LENTOS = 20

/**
 * A fake OpenAI-compatible server on 127.0.0.1, answering slowly so the test sees the text arrive.
 * `lento` keeps the stream open long enough to cancel or edit during it; `derrubar` drops the
 * connection after the first piece.
 */
export interface FakeServer {
  readonly server: Server
  readonly baseUrl: string
  readonly pedidos: string[]
  readonly corpos: string[]
  resposta: string[]
  /** With `tools` in the request and no tool result yet, the reply is these calls, in pieces. */
  chamadas: { name: string; arguments: unknown }[] | null
  fechadasCedo: number
  lento: boolean
  derrubar: boolean
  /** Answers a request with `tools` as Ollama does for a model without them. */
  semFerramentas: boolean
}

export async function startFakeServer(): Promise<FakeServer> {
  const fake = {
    pedidos: [] as string[],
    corpos: [] as string[],
    resposta: RESPOSTA,
    chamadas: null,
    fechadasCedo: 0,
    lento: false,
    derrubar: false,
    semFerramentas: false,
  } as FakeServer
  const server = createServer((req, res) => {
    fake.pedidos.push(`${req.method} ${req.url}`)
    let corpo = ''
    req.on('data', (pedaco: Buffer) => (corpo += pedaco.toString()))
    req.on('end', () => {
      if (corpo !== '') fake.corpos.push(corpo)
      if (req.url === '/v1/models') {
        res.writeHead(200, { 'content-type': 'application/json' })
        res.end(JSON.stringify({ data: [{ id: MODELO }] }))
        return
      }
      responder(res, corpo)
    })
  })
  function pedacosDe(corpo: string): string[] {
    const pedido = JSON.parse(corpo) as { tools?: unknown[]; messages: { role: string }[] }
    const primeiraRodada = pedido.tools !== undefined && !pedido.messages.some((m) => m.role === 'tool')
    if (primeiraRodada && fake.chamadas !== null) {
      return fake.chamadas.flatMap((chamada, index) => {
        const args = JSON.stringify(chamada.arguments)
        const metade = Math.floor(args.length / 2)
        return [
          {
            index,
            id: `chamada-${index}`,
            type: 'function',
            function: { name: chamada.name, arguments: args.slice(0, metade) },
          },
          { index, function: { arguments: args.slice(metade) } },
        ].map((delta) => JSON.stringify({ choices: [{ delta: { tool_calls: [delta] } }] }))
      })
    }
    const textos = fake.lento ? Array.from({ length: PEDACOS_LENTOS }, (_, i) => `${i} `) : fake.resposta
    return textos.map((texto) => JSON.stringify({ choices: [{ delta: { content: texto } }] }))
  }
  function responder(res: ServerResponse, corpo: string): void {
    if (fake.semFerramentas && (JSON.parse(corpo) as { tools?: unknown }).tools !== undefined) {
      res.writeHead(400, { 'content-type': 'application/json' })
      res.end(JSON.stringify({ error: { message: `${MODELO} does not support tools` } }))
      return
    }
    res.writeHead(200, { 'content-type': 'text/event-stream' })
    const pedacos = pedacosDe(corpo)
    let i = 0
    const timer = setInterval(() => {
      if (fake.derrubar && i === 1) {
        res.socket?.destroy()
        return
      }
      if (i < pedacos.length) {
        res.write(`data: ${pedacos[i++]}\n\n`)
        return
      }
      clearInterval(timer)
      res.end('data: [DONE]\n\n')
    }, 250)
    res.on('close', () => {
      clearInterval(timer)
      if (!res.writableEnded) fake.fechadasCedo++
    })
  }
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  const { port } = server.address() as AddressInfo
  return Object.assign(fake, { server, baseUrl: `http://127.0.0.1:${port}/v1` })
}

/** Through the setup dialog, as a person would: a server on 127.0.0.1 needs no consent. */
export async function configurarAssistente(session: Session, baseUrl: string): Promise<void> {
  await menu(session, 'ai-providers')
  const dialogo = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
  await dialogo.getByLabel('Nome').fill('Servidor falso')
  await dialogo.getByLabel('Endereço').fill(baseUrl)
  await dialogo.getByRole('button', { name: 'Salvar' }).click()
  await expect(dialogo.getByText('Servidor falso')).toBeVisible()
  await expect(dialogo.getByText('Neste computador · Padrão')).toBeVisible()
  await dialogo.getByRole('button', { name: 'Fechar' }).click()
}
