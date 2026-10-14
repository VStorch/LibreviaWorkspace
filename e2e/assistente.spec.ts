import { createServer, type Server } from 'node:http'
import type { AddressInfo } from 'node:net'
import { expect, test, type Locator } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

const MODELO = 'modelo-falso'
const RESPOSTA = ['Texto ', 'reescrito ', 'pelo ', 'modelo.']

const PEDACOS_LENTOS = 20

/**
 * A fake OpenAI-compatible server on 127.0.0.1, answering slowly so the test sees the text arrive.
 * `lento` keeps the stream open long enough to cancel or edit during it; `derrubar` drops the
 * connection after the first piece.
 */
interface FakeServer {
  readonly server: Server
  readonly baseUrl: string
  readonly pedidos: string[]
  fechadasCedo: number
  lento: boolean
  derrubar: boolean
}

async function startFakeServer(): Promise<FakeServer> {
  const fake = { pedidos: [] as string[], fechadasCedo: 0, lento: false, derrubar: false } as FakeServer
  const server = createServer((req, res) => {
    fake.pedidos.push(`${req.method} ${req.url}`)
    if (req.url === '/v1/models') {
      res.writeHead(200, { 'content-type': 'application/json' })
      res.end(JSON.stringify({ data: [{ id: MODELO }] }))
      return
    }
    res.writeHead(200, { 'content-type': 'text/event-stream' })
    const pedacos = fake.lento ? Array.from({ length: PEDACOS_LENTOS }, (_, i) => `${i} `) : RESPOSTA
    let i = 0
    const timer = setInterval(() => {
      if (fake.derrubar && i === 1) {
        res.socket?.destroy()
        return
      }
      if (i < pedacos.length) {
        res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: pedacos[i++] } }] })}\n\n`)
        return
      }
      clearInterval(timer)
      res.end('data: [DONE]\n\n')
    }, 250)
    res.on('close', () => {
      clearInterval(timer)
      if (!res.writableEnded) fake.fechadasCedo++
    })
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  const { port } = server.address() as AddressInfo
  return Object.assign(fake, { server, baseUrl: `http://127.0.0.1:${port}/v1` })
}

test.describe('assistente de IA', () => {
  let session: Session
  let fake: FakeServer
  let caixasDeMensagem: () => Promise<number>

  test.beforeEach(async () => {
    fake = await startFakeServer()
    session = await launch()
    await session.app.evaluate(({ dialog }) => {
      const global = globalThis as { mensagens?: number }
      global.mensagens = 0
      dialog.showMessageBox = async () => {
        global.mensagens = (global.mensagens ?? 0) + 1
        return { response: 1, checkboxChecked: false }
      }
    })
    caixasDeMensagem = () => session.app.evaluate(() => (globalThis as { mensagens?: number }).mensagens ?? 0)
  })

  test.afterEach(async () => {
    await session.close()
    await new Promise((resolve) => fake.server.close(resolve))
  })

  async function escrever(texto: string): Promise<Locator> {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type(texto)
    return editor
  }

  async function configurar(): Promise<void> {
    await menu(session, 'ai-providers')
    const dialogo = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    await dialogo.getByLabel('Nome').fill('Servidor falso')
    await dialogo.getByLabel('Endereço').fill(fake.baseUrl)
    await dialogo.getByRole('button', { name: 'Salvar' }).click()
    await expect(dialogo.getByText('Servidor falso')).toBeVisible()
    await expect(dialogo.getByText('Neste computador · Padrão')).toBeVisible()
    await dialogo.getByRole('button', { name: 'Fechar' }).click()
  }

  function painel() {
    return session.window.getByRole('complementary', { name: 'Assistente de IA' })
  }

  test('sem assistente configurado, nada sai do computador', async () => {
    await escrever('O gato dorme.')
    await menu(session, 'ai-assistant')
    await expect(painel().getByText('Nenhum assistente configurado')).toBeVisible()

    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await editor.click({ button: 'right', position: { x: 10, y: 5 } })
    await session.window.getByRole('menuitem', { name: 'Assistente de IA: Reescrever' }).click()
    await expect(session.window.getByRole('dialog', { name: 'Assistentes de IA' })).toBeVisible()
    expect(fake.pedidos).toEqual([])
  })

  test('reescrever chega aos poucos, entra como revisão do modelo e rejeitar devolve o original', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    const resposta = painel().getByLabel('Resposta')
    await expect(resposta).toHaveValue(/^Texto /)
    expect(await resposta.inputValue()).not.toContain('modelo.')
    await expect(resposta).toHaveValue(RESPOSTA.join(''))

    await painel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator(`ins.revision[data-author="${MODELO}"]`)).toHaveText(RESPOSTA.join(''))
    await expect(editor.locator(`del.revision[data-author="${MODELO}"]`)).toHaveText('O gato dorme.')
    expect(await caixasDeMensagem()).toBe(0)
    expect(fake.pedidos).toContain('POST /v1/chat/completions')

    await menu(session, 'reject-all-changes')
    await expect(editor.locator('ins.revision, del.revision')).toHaveCount(0)
    await expect(editor).toHaveText('O gato dorme.')
  })

  test('parar fecha a conexão com o assistente', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    fake.lento = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveValue(/^0 /)
    await painel().getByRole('button', { name: 'Parar' }).click()

    await expect(painel().getByText('Interrompido.')).toBeVisible()
    await expect.poll(() => fake.fechadasCedo).toBe(1)
  })

  test('editar o trecho durante a resposta desativa Aplicar', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    fake.lento = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveValue(/^0 /)
    await editor.click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.type('X')

    await expect(painel().getByText('O trecho mudou')).toBeVisible()
    await expect(painel().getByRole('button', { name: 'Inserir abaixo' })).toBeEnabled({ timeout: 10_000 })
    await expect(painel().getByRole('button', { name: 'Aplicar' })).toBeDisabled()
  })

  test('resposta interrompida não pode ser aplicada', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    fake.lento = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveValue(/^0 /)
    await painel().getByRole('button', { name: 'Parar' }).click()

    await expect(painel().getByText('Interrompido.')).toBeVisible()
    await expect(painel().getByRole('button', { name: 'Aplicar' })).toBeDisabled()
    await expect(painel().getByRole('button', { name: 'Inserir abaixo' })).toBeDisabled()
  })

  test('provedor em 127.0.0.1 não pede consentimento e busca a lista de modelos uma vez só', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    const gerar = painel().getByRole('button', { name: 'Gerar' })
    const resposta = painel().getByLabel('Resposta')
    await gerar.click()
    await expect(resposta).toHaveValue(RESPOSTA.join(''))
    await gerar.click()
    await expect(resposta).toHaveValue(/^Texto /)
    await expect(resposta).toHaveValue(RESPOSTA.join(''))

    expect(await caixasDeMensagem()).toBe(0)
    expect(fake.pedidos.filter((pedido) => pedido === 'POST /v1/chat/completions')).toHaveLength(2)
    expect(fake.pedidos.filter((pedido) => pedido === 'GET /v1/models')).toHaveLength(1)
  })

  test('conexão derrubada no meio da resposta dá aviso de rede, não erro genérico', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    fake.derrubar = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(session.window.getByText('Não foi possível falar com o assistente de IA.')).toBeVisible()
    await expect(session.window.getByText('Ocorreu um erro inesperado.')).toHaveCount(0)
    await expect(painel().getByRole('button', { name: 'Aplicar' })).toBeDisabled()
  })
})
