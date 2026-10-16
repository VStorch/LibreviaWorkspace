import { existsSync } from 'node:fs'
import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { createServer, type Server } from 'node:http'
import type { AddressInfo } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Locator } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithCellRevision } from './fixtures.js'

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
  readonly corpos: string[]
  resposta: string[]
  fechadasCedo: number
  lento: boolean
  derrubar: boolean
}

async function startFakeServer(): Promise<FakeServer> {
  const fake = {
    pedidos: [] as string[],
    corpos: [] as string[],
    resposta: RESPOSTA,
    fechadasCedo: 0,
    lento: false,
    derrubar: false,
  } as FakeServer
  const server = createServer((req, res) => {
    fake.pedidos.push(`${req.method} ${req.url}`)
    let corpo = ''
    req.on('data', (pedaco: Buffer) => (corpo += pedaco.toString()))
    req.on('end', () => {
      if (corpo !== '') fake.corpos.push(corpo)
    })
    if (req.url === '/v1/models') {
      res.writeHead(200, { 'content-type': 'application/json' })
      res.end(JSON.stringify({ data: [{ id: MODELO }] }))
      return
    }
    res.writeHead(200, { 'content-type': 'text/event-stream' })
    const pedacos = fake.lento ? Array.from({ length: PEDACOS_LENTOS }, (_, i) => `${i} `) : fake.resposta
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

  async function acaoDoMenu(acao: string): Promise<void> {
    await session.window.getByRole('menuitem', { name: 'Assistente de IA' }).hover()
    await session.window
      .getByRole('menu', { name: 'Assistente de IA' })
      .getByRole('menuitem', { name: acao })
      .click()
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
    await acaoDoMenu('Reescrever')
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
    await expect(resposta).toHaveText(/^Texto /)
    expect(await resposta.textContent()).not.toContain('modelo.')
    await expect(resposta).toHaveText(RESPOSTA.join(''))

    await painel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator(`ins.revision[data-author="${MODELO}"]`)).toHaveText(RESPOSTA.join(''))
    await expect(editor.locator(`del.revision[data-author="${MODELO}"]`)).toHaveText('O gato dorme.')
    expect(await caixasDeMensagem()).toBe(0)
    expect(fake.pedidos).toContain('POST /v1/chat/completions')

    await menu(session, 'reject-all-changes')
    await expect(editor.locator('ins.revision, del.revision')).toHaveCount(0)
    await expect(editor).toHaveText('O gato dorme.')
  })

  test('o mesmo comando abre e fecha o painel, e reabrir não repete o pedido', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await editor.click({ button: 'right', position: { x: 10, y: 5 } })
    await acaoDoMenu('Reescrever')
    await expect(painel().getByLabel('Resposta')).toHaveText(RESPOSTA.join(''))
    const gerados = () => fake.pedidos.filter((pedido) => pedido.startsWith('POST')).length
    expect(gerados()).toBe(1)

    await menu(session, 'ai-assistant')
    await expect(painel()).toHaveCount(0)
    await menu(session, 'ai-assistant')
    await expect(painel()).toBeVisible()
    await expect(painel().getByRole('button', { name: 'Gerar' })).toBeEnabled()
    expect(gerados()).toBe(1)
  })

  test('parar fecha a conexão com o assistente', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    fake.lento = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveText(/^0 /)
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
    await expect(painel().getByLabel('Resposta')).toHaveText(/^0 /)
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
    await expect(painel().getByLabel('Resposta')).toHaveText(/^0 /)
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
    await expect(resposta).toHaveText(RESPOSTA.join(''))
    await gerar.click()
    await expect(resposta).toHaveText(/^Texto /)
    await expect(resposta).toHaveText(RESPOSTA.join(''))

    expect(await caixasDeMensagem()).toBe(0)
    expect(fake.pedidos.filter((pedido) => pedido === 'POST /v1/chat/completions')).toHaveLength(2)
    expect(fake.pedidos.filter((pedido) => pedido === 'GET /v1/models')).toHaveLength(1)
  })

  test('o modelo escolhido no painel fica lembrado ao reabrir', async () => {
    await escrever('O gato dorme.')
    await configurar()
    await menu(session, 'ai-assistant')
    const modelo = painel().getByLabel('Modelo')
    await expect(modelo).toHaveValue('')
    await painel().getByRole('button', { name: 'Listar modelos' }).click()
    await modelo.selectOption(MODELO)
    await expect(modelo).toHaveValue(MODELO)

    await menu(session, 'ai-assistant')
    await expect(painel()).toHaveCount(0)
    await menu(session, 'ai-assistant')
    await expect(painel().getByLabel('Modelo')).toHaveValue(MODELO)
  })

  test('editar o assistente troca o nome e mantém o endereço', async () => {
    await escrever('O gato dorme.')
    await configurar()
    await menu(session, 'ai-providers')
    const dialogo = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    await dialogo.getByRole('button', { name: 'Editar' }).click()
    const edicao = dialogo.getByRole('group', { name: 'Editar assistente' })
    await expect(edicao.getByLabel('Endereço')).toHaveValue(fake.baseUrl)
    await edicao.getByLabel('Nome').fill('Servidor renomeado')
    await edicao.getByRole('button', { name: 'Salvar' }).click()
    await expect(dialogo.getByText('Servidor renomeado')).toBeVisible()
    await expect(dialogo.getByText('Servidor falso')).toHaveCount(0)
    await expect(dialogo.getByText('Neste computador · Padrão')).toBeVisible()
  })

  test('o submenu do assistente funciona pelo teclado, e Esc fecha só o submenu', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await session.window.keyboard.press('Shift+F10')
    const contexto = session.window.getByRole('menu', { name: 'Ações do documento' })
    const gatilho = contexto.getByRole('menuitem', { name: 'Assistente de IA' })
    await expect(contexto.getByRole('menuitem').first()).toBeFocused()

    await session.window.keyboard.press('End')
    await expect(gatilho).toBeFocused()
    await session.window.keyboard.press('ArrowRight')
    const submenu = session.window.getByRole('menu', { name: 'Assistente de IA' })
    await expect(submenu.getByRole('menuitem', { name: 'Reescrever' })).toBeFocused()

    await session.window.keyboard.press('Escape')
    await expect(submenu).toHaveCount(0)
    await expect(contexto).toBeVisible()
    await expect(gatilho).toBeFocused()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('ArrowDown')
    await session.window.keyboard.press('ArrowDown')
    await expect(submenu.getByRole('menuitem', { name: 'Revisar ortografia e gramática' })).toBeFocused()
    await session.window.keyboard.press('Enter')
    await expect(contexto).toHaveCount(0)
    await expect(painel().getByLabel('Resposta')).toHaveText(RESPOSTA.join(''))
  })

  function mensagensEnviadas(): { system: string; user: string } {
    const corpo = JSON.parse(fake.corpos.at(-1) ?? '{}') as { messages?: { role: string; content: string }[] }
    const conteudo = (papel: string) => corpo.messages?.find((m) => m.role === papel)?.content ?? ''
    return { system: conteudo('system'), user: conteudo('user') }
  }

  async function selecionarParagrafoDoMeio(): Promise<Locator> {
    const editor = await escrever('Primeiro parágrafo.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('O gato dorme.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Último parágrafo.')
    await configurar()
    await editor.getByText('O gato dorme.').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Shift+End')
    return editor
  }

  test('o documento vai junto como contexto, com a seleção marcada', async () => {
    await selecionarParagrafoDoMeio()
    await menu(session, 'ai-assistant')
    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveText(RESPOSTA.join(''))

    const { user } = mensagensEnviadas()
    expect(user).toContain('Primeiro parágrafo.')
    expect(user).toContain('<selection>O gato dorme.</selection>')
    expect(user).toContain('Último parágrafo.')
  })

  test('resposta em Markdown aparece formatada e, aplicada, vira negrito e lista na revisão do modelo', async () => {
    const editor = await selecionarParagrafoDoMeio()
    fake.resposta = ['**Gato** ', 'dorme.\n\n', '- item um\n', '- item dois']
    await menu(session, 'ai-assistant')
    await painel().getByRole('button', { name: 'Gerar' }).click()
    const resposta = painel().getByLabel('Resposta')
    await expect(resposta.locator('li')).toHaveText(['item um', 'item dois'])
    await expect(resposta.locator('strong')).toHaveText('Gato')

    await painel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator('strong', { hasText: 'Gato' })).toHaveCount(1)
    await expect(editor.locator('li')).toHaveText(['item um', 'item dois'])
    await expect(editor.locator(`ins.revision[data-author="${MODELO}"]`, { hasText: 'item um' })).toHaveCount(
      1,
    )
    await expect(editor.locator(`del.revision[data-author="${MODELO}"]`)).toHaveText('O gato dorme.')
  })

  test('sem manter a formatação, a revisão vai e volta em texto puro', async () => {
    await selecionarParagrafoDoMeio()
    await menu(session, 'ai-providers')
    const dialogo = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    const manter = dialogo.getByLabel('Manter a formatação nas revisões')
    await manter.click()
    await expect(manter).not.toBeChecked()
    await dialogo.getByRole('button', { name: 'Fechar' }).click()
    await menu(session, 'ai-assistant')
    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveValue(RESPOSTA.join(''))

    const { system } = mensagensEnviadas()
    expect(system).toContain('no Markdown')
    expect(system).not.toContain('Format it in Markdown')
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

  async function abrirConversa(): Promise<{ campo: Locator; conversa: Locator }> {
    await menu(session, 'ai-assistant')
    await painel().getByRole('tab', { name: 'Conversa' }).click()
    return { campo: painel().getByLabel('Mensagem'), conversa: painel().getByRole('log') }
  }

  async function perguntar(campo: Locator, pergunta: string): Promise<void> {
    await campo.fill(pergunta)
    await campo.press('Enter')
  }

  function conversaEnviada(): { role: string; content: string }[] {
    const corpo = JSON.parse(fake.corpos.at(-1) ?? '{}') as { messages?: { role: string; content: string }[] }
    return corpo.messages ?? []
  }

  test('na conversa, o segundo pedido leva a primeira pergunta e a primeira resposta, e Nova conversa limpa', async () => {
    await escrever('O gato dorme.')
    await configurar()
    const { campo, conversa } = await abrirConversa()

    await perguntar(campo, 'Qual é o tema?')
    await expect(campo).toBeFocused()
    await expect(campo).toHaveValue('')
    await expect(conversa.getByLabel('Resposta')).toHaveText(RESPOSTA.join(''))

    await perguntar(campo, 'E o tom?')
    await expect(conversa.getByLabel('Resposta')).toHaveCount(2)
    await expect(conversa.getByLabel('Resposta').nth(1)).toHaveText(RESPOSTA.join(''))
    const [system, ...historico] = conversaEnviada()
    expect(system?.role).toBe('system')
    expect(system?.content).toContain('O gato dorme.')
    expect(historico).toEqual([
      { role: 'user', content: 'Qual é o tema?' },
      { role: 'assistant', content: RESPOSTA.join('') },
      { role: 'user', content: 'E o tom?' },
    ])

    await painel().getByRole('tab', { name: 'Ações' }).click()
    await expect(painel().getByRole('button', { name: 'Gerar' })).toBeVisible()
    await painel().getByRole('tab', { name: 'Conversa' }).click()
    await expect(conversa.getByText('E o tom?')).toBeVisible()

    await painel().getByRole('button', { name: 'Nova conversa' }).click()
    await expect(conversa.getByLabel('Resposta')).toHaveCount(0)
    await expect(conversa.getByText('Qual é o tema?')).toHaveCount(0)
    await expect(campo).toBeFocused()
  })

  test('Inserir abaixo da conversa entra como revisão com o modelo como autor', async () => {
    const editor = await escrever('O gato dorme.')
    await configurar()
    const { campo, conversa } = await abrirConversa()
    await perguntar(campo, 'Escreva uma frase.')
    await expect(conversa.getByLabel('Resposta')).toHaveText(RESPOSTA.join(''))

    await conversa.getByRole('button', { name: 'Inserir abaixo' }).click()
    await expect(editor.locator(`ins.revision[data-author="${MODELO}"]`)).toHaveText(RESPOSTA.join(''))
    await expect(editor.getByText('O gato dorme.')).toBeVisible()
  })

  test('documento só para leitura: o painel abre, a conversa funciona e Inserir abaixo fica desativado', async () => {
    const pasta = await mkdtemp(join(tmpdir(), 'librevia-ia-leitura-'))
    try {
      const origem = join(pasta, 'travado.docx')
      await writeFile(origem, await docxWithCellRevision())
      await stubDialogs(session.app, { open: origem, messageBox: 1 })
      await menu(session, 'open')
      await expect(session.window.locator('.banner--readonly')).toBeVisible()
      await configurar()

      const { campo, conversa } = await abrirConversa()
      await perguntar(campo, 'Do que trata?')
      await expect(conversa.getByLabel('Resposta')).toHaveText(RESPOSTA.join(''))
      await expect(conversa.getByRole('button', { name: 'Inserir abaixo' })).toBeDisabled()
    } finally {
      await rm(pasta, { recursive: true, force: true })
    }
  })

  test('parar no meio da conversa fecha a conexão e deixa enviar de novo', async () => {
    await escrever('O gato dorme.')
    await configurar()
    fake.lento = true
    const { campo, conversa } = await abrirConversa()
    await perguntar(campo, 'Conte devagar.')
    await expect(conversa.getByLabel('Resposta')).toHaveText(/^0 /)

    await painel().getByRole('button', { name: 'Parar' }).click()
    await expect(conversa.getByText('Interrompido.')).toBeVisible()
    await expect.poll(() => fake.fechadasCedo).toBe(1)
    await expect(conversa.getByRole('button', { name: 'Inserir abaixo' })).toHaveCount(0)
    await campo.fill('De novo')
    await expect(painel().getByRole('button', { name: 'Enviar' })).toBeEnabled()
  })
})

/**
 * A fake ACP agent, JSON-RPC written by hand: it thinks, calls a tool, asks permission, and answers
 * only after the permission reply. It logs every message it receives, its pid and its folder. It
 * ignores `session/cancel`, so stopping must end the process. "devagar" in the prompt makes it slow.
 */
function fakeAgentSource(logPath: string): string {
  return `#!/usr/bin/env node
import { appendFileSync } from 'node:fs'
const log = (entry) => appendFileSync(${JSON.stringify(logPath)}, JSON.stringify(entry) + '\\n')
const send = (message) => process.stdout.write(JSON.stringify({ jsonrpc: '2.0', ...message }) + '\\n')
const update = (update) => send({ method: 'session/update', params: { sessionId: 's1', update } })
const MODEL = {
  id: 'model', name: 'Modelo', type: 'select', category: 'model', currentValue: 'agente-rapido',
  options: [{ value: 'agente-rapido', name: 'Rápido' }, { value: 'agente-esperto', name: 'Esperto' }],
}
log({ pid: process.pid, cwd: process.cwd() })
let prompt
let buffer = ''
process.stdin.on('data', (chunk) => {
  buffer += chunk
  let end
  while ((end = buffer.indexOf('\\n')) >= 0) {
    const line = buffer.slice(0, end)
    buffer = buffer.slice(end + 1)
    if (line.trim() !== '') handle(JSON.parse(line))
  }
})
function handle(message) {
  log(message)
  if (message.method === 'initialize') return send({ id: message.id, result: { protocolVersion: 1, agentCapabilities: {} } })
  if (message.method === 'session/new') return send({ id: message.id, result: { sessionId: 's1', configOptions: [MODEL] } })
  if (message.method === 'session/set_config_option') return send({ id: message.id, result: { configOptions: [MODEL] } })
  if (message.method === 'session/prompt') {
    prompt = message
    update({ sessionUpdate: 'agent_thought_chunk', content: { type: 'text', text: 'pensamento secreto' } })
    update({ sessionUpdate: 'tool_call', toolCallId: 't1', title: 'Apagar arquivos', kind: 'delete' })
    return send({
      id: 'permissao',
      method: 'session/request_permission',
      params: {
        sessionId: 's1',
        toolCall: { toolCallId: 't1' },
        options: [
          { optionId: 'permitir', name: 'Permitir', kind: 'allow_always' },
          { optionId: 'recusar', name: 'Recusar', kind: 'reject_once' },
        ],
      },
    })
  }
  if (message.id === 'permissao') answer(JSON.stringify(prompt.params.prompt).includes('devagar'))
}
function answer(slow) {
  const pieces = slow ? Array.from({ length: 200 }, (_, i) => i + ' ') : ['Texto ', 'do ', 'agente.']
  let next = 0
  const timer = setInterval(() => {
    if (next < pieces.length) {
      update({ sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: pieces[next++] } })
      return
    }
    clearInterval(timer)
    send({ id: prompt.id, result: { stopReason: 'end_turn' } })
  }, 200)
}
`
}

interface AgentLogEntry {
  readonly pid?: number
  readonly cwd?: string
  readonly id?: string | number
  readonly method?: string
  readonly params?: Record<string, unknown>
  readonly result?: unknown
}

test.describe('assistente por agente ACP', () => {
  test.skip(process.platform === 'win32', 'o agente falso é um script com shebang')

  let session: Session
  let pasta: string
  let registro: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-acp-e2e-'))
    registro = join(pasta, 'registro.jsonl')
    const agente = join(pasta, 'fake-acp-agent.mjs')
    await writeFile(agente, fakeAgentSource(registro))
    await chmod(agente, 0o755)
    session = await launch()
    await stubDialogs(session.app, { open: agente, messageBox: 0 })
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  async function lerRegistro(): Promise<AgentLogEntry[]> {
    const texto = await readFile(registro, 'utf8').catch(() => '')
    return texto
      .split('\n')
      .filter((linha) => linha !== '')
      .map((linha) => JSON.parse(linha) as AgentLogEntry)
  }

  function painel() {
    return session.window.getByRole('complementary', { name: 'Assistente de IA' })
  }

  async function prepararAgente(texto: string): Promise<Locator> {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type(texto)
    await menu(session, 'ai-providers')
    const dialogo = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    await dialogo.getByLabel('Tipo').selectOption({ label: 'Agente ACP (programa instalado)' })
    await dialogo.getByLabel('Nome').fill('Agente falso')
    await dialogo.getByRole('button', { name: 'Escolher o programa…' }).click()
    await expect(dialogo.getByText('Agente falso')).toBeVisible()
    await expect(dialogo.getByText(/Fora deste computador · fake-acp-agent\.mjs/)).toBeVisible()
    await dialogo.getByRole('button', { name: 'Fechar' }).click()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')
    return editor
  }

  test('o agente responde só com texto, a permissão é recusada e o modelo escolhido vira o autor', async () => {
    const editor = await prepararAgente('O gato dorme.')
    await painel().getByRole('button', { name: 'Listar modelos' }).click()
    const modelo = painel().getByLabel('Modelo')
    await modelo.selectOption('agente-esperto')

    await painel().getByRole('button', { name: 'Gerar' }).click()
    const resposta = painel().getByLabel('Resposta')
    await expect(resposta).toHaveText('Texto do agente.')
    await expect(resposta).not.toContainText('pensamento')

    await painel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator('ins.revision[data-author="agente-esperto"]')).toHaveText('Texto do agente.')

    const mensagens = await lerRegistro()
    const permissao = mensagens.find((mensagem) => mensagem.id === 'permissao')
    expect(permissao?.result).toEqual({ outcome: { outcome: 'selected', optionId: 'recusar' } })
    const escolha = mensagens.find((mensagem) => mensagem.method === 'session/set_config_option')
    expect(escolha?.params).toMatchObject({ configId: 'model', value: 'agente-esperto' })
    const inicio = mensagens.find((mensagem) => mensagem.method === 'initialize')
    expect(inicio?.params?.['clientCapabilities']).toEqual({
      fs: { readTextFile: false, writeTextFile: false },
      terminal: false,
    })
    const pastas = mensagens.flatMap((mensagem) => (mensagem.cwd === undefined ? [] : [mensagem.cwd]))
    expect(pastas.every((cwd) => cwd.includes('librevia-ai-'))).toBe(true)
    await expect.poll(() => pastas.some((cwd) => existsSync(cwd))).toBe(false)
  })

  test('parar avisa o agente e encerra o processo dele', async () => {
    await prepararAgente('Escreva devagar.')
    await painel().getByRole('button', { name: 'Gerar' }).click()
    await expect(painel().getByLabel('Resposta')).toHaveText(/^0 /)
    await painel().getByRole('button', { name: 'Parar' }).click()
    await expect(painel().getByText('Interrompido.')).toBeVisible()

    await expect
      .poll(async () => (await lerRegistro()).some((mensagem) => mensagem.method === 'session/cancel'))
      .toBe(true)
    const mensagens = await lerRegistro()
    const pids = mensagens.flatMap((mensagem) => (mensagem.pid === undefined ? [] : [mensagem.pid]))
    await expect.poll(() => pids.filter(estaVivo)).toEqual([])
  })
})

function estaVivo(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}
