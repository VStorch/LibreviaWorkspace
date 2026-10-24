import { existsSync } from 'node:fs'
import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Locator } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { configureAssistant, TEMPLATE, REPLY, startFakeServer, type FakeServer } from './fake-ai-server.js'
import { docxWithCellRevision } from './fixtures.js'

test.describe('AI assistant', () => {
  let session: Session
  let fake: FakeServer
  let messageBoxes: () => Promise<number>

  test.beforeEach(async () => {
    fake = await startFakeServer()
    session = await launch()
    await session.app.evaluate(({ dialog }) => {
      const global = globalThis as { chatMessages?: number }
      global.chatMessages = 0
      dialog.showMessageBox = async () => {
        global.chatMessages = (global.chatMessages ?? 0) + 1
        return { response: 1, checkboxChecked: false }
      }
    })
    messageBoxes = () =>
      session.app.evaluate(() => (globalThis as { chatMessages?: number }).chatMessages ?? 0)
  })

  test.afterEach(async () => {
    await session.close()
    await new Promise((resolve) => fake.server.close(resolve))
  })

  async function write(text: string): Promise<Locator> {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type(text)
    return editor
  }

  async function configure(): Promise<void> {
    await configureAssistant(session, fake.baseUrl)
  }

  async function menuAction(action: string): Promise<void> {
    await session.window.getByRole('menuitem', { name: 'Assistente de IA' }).hover()
    await session.window
      .getByRole('menu', { name: 'Assistente de IA' })
      .getByRole('menuitem', { name: action })
      .click()
  }

  function panel() {
    return session.window.getByRole('complementary', { name: 'Assistente de IA' })
  }

  test('with no assistant configured, nothing leaves the computer', async () => {
    await write('O gato dorme.')
    await menu(session, 'ai-assistant')
    await expect(panel().getByText('Nenhum assistente configurado')).toBeVisible()

    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await editor.click({ button: 'right', position: { x: 10, y: 5 } })
    await menuAction('Reescrever')
    await expect(session.window.getByRole('dialog', { name: 'Assistentes de IA' })).toBeVisible()
    expect(fake.requests).toEqual([])
  })

  test('rewrite arrives gradually, enters as a model revision, and rejecting restores the original', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await panel().getByRole('button', { name: 'Gerar' }).click()
    const reply = panel().getByLabel('Resposta')
    await expect(reply).toHaveText(/^Texto /)
    expect(await reply.textContent()).not.toContain('modelo.')
    await expect(reply).toHaveText(REPLY.join(''))

    await panel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator(`ins.revision[data-author="${TEMPLATE}"]`)).toHaveText(REPLY.join(''))
    await expect(editor.locator(`del.revision[data-author="${TEMPLATE}"]`)).toHaveText('O gato dorme.')
    expect(await messageBoxes()).toBe(0)
    expect(fake.requests).toContain('POST /v1/chat/completions')

    await menu(session, 'reject-all-changes')
    await expect(editor.locator('ins.revision, del.revision')).toHaveCount(0)
    await expect(editor).toHaveText('O gato dorme.')
  })

  test('the same command opens and closes the panel, and reopening does not repeat the request', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await editor.click({ button: 'right', position: { x: 10, y: 5 } })
    await menuAction('Reescrever')
    await expect(panel().getByLabel('Resposta')).toHaveText(REPLY.join(''))
    const generated = () => fake.requests.filter((request) => request.startsWith('POST')).length
    expect(generated()).toBe(1)

    await menu(session, 'ai-assistant')
    await expect(panel()).toHaveCount(0)
    await menu(session, 'ai-assistant')
    await expect(panel()).toBeVisible()
    await expect(panel().getByRole('button', { name: 'Gerar' })).toBeEnabled()
    expect(generated()).toBe(1)
  })

  test('stop closes the connection to the assistant', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    fake.slow = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(panel().getByLabel('Resposta')).toHaveText(/^0 /)
    await panel().getByRole('button', { name: 'Parar' }).click()

    await expect(panel().getByText('Interrompido.')).toBeVisible()
    await expect.poll(() => fake.closedEarly).toBe(1)
  })

  test('editing the excerpt during the reply disables Apply', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    fake.slow = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(panel().getByLabel('Resposta')).toHaveText(/^0 /)
    await editor.click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.type('X')

    await expect(panel().getByText('O trecho mudou')).toBeVisible()
    await expect(panel().getByRole('button', { name: 'Inserir abaixo' })).toBeEnabled({ timeout: 10_000 })
    await expect(panel().getByRole('button', { name: 'Aplicar' })).toBeDisabled()
  })

  test('an interrupted reply cannot be applied', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    fake.slow = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(panel().getByLabel('Resposta')).toHaveText(/^0 /)
    await panel().getByRole('button', { name: 'Parar' }).click()

    await expect(panel().getByText('Interrompido.')).toBeVisible()
    await expect(panel().getByRole('button', { name: 'Aplicar' })).toBeDisabled()
    await expect(panel().getByRole('button', { name: 'Inserir abaixo' })).toBeDisabled()
  })

  test('a provider on 127.0.0.1 asks for no consent and fetches the model list only once', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    const generate = panel().getByRole('button', { name: 'Gerar' })
    const reply = panel().getByLabel('Resposta')
    await generate.click()
    await expect(reply).toHaveText(REPLY.join(''))
    await generate.click()
    await expect(reply).toHaveText(/^Texto /)
    await expect(reply).toHaveText(REPLY.join(''))

    expect(await messageBoxes()).toBe(0)
    expect(fake.requests.filter((request) => request === 'POST /v1/chat/completions')).toHaveLength(2)
    expect(fake.requests.filter((request) => request === 'GET /v1/models')).toHaveLength(1)
  })

  test('the model chosen in the panel is remembered on reopening', async () => {
    await write('O gato dorme.')
    await configure()
    await menu(session, 'ai-assistant')
    const template = panel().getByLabel('Modelo')
    await expect(template).toHaveValue('')
    await panel().getByRole('button', { name: 'Listar modelos' }).click()
    await template.selectOption(TEMPLATE)
    await expect(template).toHaveValue(TEMPLATE)

    await menu(session, 'ai-assistant')
    await expect(panel()).toHaveCount(0)
    await menu(session, 'ai-assistant')
    await expect(panel().getByLabel('Modelo')).toHaveValue(TEMPLATE)
  })

  test('editing the assistant changes the name and keeps the address', async () => {
    await write('O gato dorme.')
    await configure()
    await menu(session, 'ai-providers')
    const dialog = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    await dialog.getByRole('button', { name: 'Editar' }).click()
    const editing = dialog.getByRole('group', { name: 'Editar assistente' })
    await expect(editing.getByLabel('Endereço')).toHaveValue(fake.baseUrl)
    await editing.getByLabel('Nome').fill('Servidor renomeado')
    await editing.getByRole('button', { name: 'Salvar' }).click()
    await expect(dialog.getByText('Servidor renomeado')).toBeVisible()
    await expect(dialog.getByText('Servidor falso')).toHaveCount(0)
    await expect(dialog.getByText('Neste computador · Padrão')).toBeVisible()
  })

  test('the assistant submenu works by keyboard, and Esc closes only the submenu', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await session.window.keyboard.press('Shift+F10')
    const context = session.window.getByRole('menu', { name: 'Ações do documento' })
    const trigger = context.getByRole('menuitem', { name: 'Assistente de IA' })
    await expect(context.getByRole('menuitem').first()).toBeFocused()

    await session.window.keyboard.press('End')
    await expect(trigger).toBeFocused()
    await session.window.keyboard.press('ArrowRight')
    const submenu = session.window.getByRole('menu', { name: 'Assistente de IA' })
    await expect(submenu.getByRole('menuitem', { name: 'Reescrever' })).toBeFocused()

    await session.window.keyboard.press('Escape')
    await expect(submenu).toHaveCount(0)
    await expect(context).toBeVisible()
    await expect(trigger).toBeFocused()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('ArrowDown')
    await session.window.keyboard.press('ArrowDown')
    await expect(submenu.getByRole('menuitem', { name: 'Revisar ortografia e gramática' })).toBeFocused()
    await session.window.keyboard.press('Enter')
    await expect(context).toHaveCount(0)
    await expect(panel().getByLabel('Resposta')).toHaveText(REPLY.join(''))
  })

  function sentMessages(): { system: string; user: string } {
    const body = JSON.parse(fake.bodies.at(-1) ?? '{}') as { messages?: { role: string; content: string }[] }
    const contents = (paper: string) => body.messages?.find((m) => m.role === paper)?.content ?? ''
    return { system: contents('system'), user: contents('user') }
  }

  async function selectMiddleParagraph(): Promise<Locator> {
    const editor = await write('Primeiro parágrafo.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('O gato dorme.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Último parágrafo.')
    await configure()
    await editor.getByText('O gato dorme.').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Shift+End')
    return editor
  }

  test('the document goes along as context, with the selection marked', async () => {
    await selectMiddleParagraph()
    await menu(session, 'ai-assistant')
    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(panel().getByLabel('Resposta')).toHaveText(REPLY.join(''))

    const { user } = sentMessages()
    expect(user).toContain('Primeiro parágrafo.')
    expect(user).toContain('<selection>O gato dorme.</selection>')
    expect(user).toContain('Último parágrafo.')
  })

  test('a Markdown reply shows formatted and, once applied, becomes bold and a list in the model revision', async () => {
    const editor = await selectMiddleParagraph()
    fake.reply = ['**Gato** ', 'dorme.\n\n', '- item um\n', '- item dois']
    await menu(session, 'ai-assistant')
    await panel().getByRole('button', { name: 'Gerar' }).click()
    const reply = panel().getByLabel('Resposta')
    await expect(reply.locator('li')).toHaveText(['item um', 'item dois'])
    await expect(reply.locator('strong')).toHaveText('Gato')

    await panel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator('strong', { hasText: 'Gato' })).toHaveCount(1)
    await expect(editor.locator('li')).toHaveText(['item um', 'item dois'])
    await expect(
      editor.locator(`ins.revision[data-author="${TEMPLATE}"]`, { hasText: 'item um' }),
    ).toHaveCount(1)
    await expect(editor.locator(`del.revision[data-author="${TEMPLATE}"]`)).toHaveText('O gato dorme.')
  })

  test('without keeping formatting, the revision goes and comes back as plain text', async () => {
    await selectMiddleParagraph()
    await menu(session, 'ai-providers')
    const dialog = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    const keep = dialog.getByLabel('Manter a formatação nas revisões')
    await keep.click()
    await expect(keep).not.toBeChecked()
    await dialog.getByRole('button', { name: 'Fechar' }).click()
    await menu(session, 'ai-assistant')
    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(panel().getByLabel('Resposta')).toHaveValue(REPLY.join(''))

    const { system } = sentMessages()
    expect(system).toContain('no Markdown')
    expect(system).not.toContain('Format it in Markdown')
  })

  test('a connection dropped mid-reply gives a network notice, not a generic error', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    fake.dropConnection = true
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')

    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(session.window.getByText('Não foi possível falar com o assistente de IA.')).toBeVisible()
    await expect(session.window.getByText('Ocorreu um erro inesperado.')).toHaveCount(0)
    await expect(panel().getByRole('button', { name: 'Aplicar' })).toBeDisabled()
  })

  async function openConversation(): Promise<{ field: Locator; conversation: Locator }> {
    await menu(session, 'ai-assistant')
    await panel().getByRole('tab', { name: 'Conversa' }).click()
    return { field: panel().getByLabel('Mensagem'), conversation: panel().getByRole('log') }
  }

  async function askQuestion(field: Locator, question: string): Promise<void> {
    await field.fill(question)
    await field.press('Enter')
  }

  function sentConversation(): { role: string; content: string }[] {
    const body = JSON.parse(fake.bodies.at(-1) ?? '{}') as { messages?: { role: string; content: string }[] }
    return body.messages ?? []
  }

  test('in the conversation, the second request carries the first question and first reply, and New conversation clears', async () => {
    await write('O gato dorme.')
    await configure()
    const { field, conversation } = await openConversation()

    await askQuestion(field, 'Qual é o tema?')
    await expect(field).toBeFocused()
    await expect(field).toHaveValue('')
    await expect(conversation.getByLabel('Resposta')).toHaveText(REPLY.join(''))

    await askQuestion(field, 'E o tom?')
    await expect(conversation.getByLabel('Resposta')).toHaveCount(2)
    await expect(conversation.getByLabel('Resposta').nth(1)).toHaveText(REPLY.join(''))
    const [system, ...history] = sentConversation()
    expect(system?.role).toBe('system')
    expect(system?.content).toContain('O gato dorme.')
    expect(history).toEqual([
      { role: 'user', content: 'Qual é o tema?' },
      { role: 'assistant', content: REPLY.join('') },
      { role: 'user', content: 'E o tom?' },
    ])

    await panel().getByRole('tab', { name: 'Ações' }).click()
    await expect(panel().getByRole('button', { name: 'Gerar' })).toBeVisible()
    await panel().getByRole('tab', { name: 'Conversa' }).click()
    await expect(conversation.getByText('E o tom?')).toBeVisible()

    await panel().getByRole('button', { name: 'Nova conversa' }).click()
    await expect(conversation.getByLabel('Resposta')).toHaveCount(0)
    await expect(conversation.getByText('Qual é o tema?')).toHaveCount(0)
    await expect(field).toBeFocused()
  })

  test('Insert below from the conversation enters as a revision with the model as author', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    const { field, conversation } = await openConversation()
    await askQuestion(field, 'Escreva uma frase.')
    await expect(conversation.getByLabel('Resposta')).toHaveText(REPLY.join(''))

    await conversation.getByRole('button', { name: 'Inserir abaixo' }).click()
    await expect(editor.locator(`ins.revision[data-author="${TEMPLATE}"]`)).toHaveText(REPLY.join(''))
    await expect(editor.getByText('O gato dorme.')).toBeVisible()
  })

  test('read-only document: the panel opens, the conversation works and Insert below is disabled', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'librevia-ai-readonly-'))
    try {
      const sourceFile = join(directory, 'travado.docx')
      await writeFile(sourceFile, await docxWithCellRevision())
      await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
      await menu(session, 'open')
      await expect(session.window.locator('.banner--readonly')).toBeVisible()
      await configure()

      const { field, conversation } = await openConversation()
      await askQuestion(field, 'Do que trata?')
      await expect(conversation.getByLabel('Resposta')).toHaveText(REPLY.join(''))
      await expect(conversation.getByRole('button', { name: 'Inserir abaixo' })).toBeDisabled()
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })

  const PROPOSE = 'Deixar o assistente propor alterações e comentários'

  test('with the box checked, proposals appear and Apply all creates the model revision and comment', async () => {
    const editor = await write('O gato dorme.')
    await configure()
    fake.calls = [
      { name: 'propose_edit', arguments: { paragraph: 1, find: 'dorme', replace: 'cochila', reason: 'Tom' } },
      { name: 'add_comment', arguments: { paragraph: 1, find: 'gato', comment: 'Qual gato?' } },
    ]
    const { field, conversation } = await openConversation()
    const box = panel().getByLabel(PROPOSE)
    await expect(box).not.toBeChecked()
    await box.click()
    await expect(box).toBeChecked()
    await askQuestion(field, 'Revise o texto.')

    await expect(conversation.getByLabel('Resposta')).toHaveText(REPLY.join(''))
    const proposals = conversation.getByRole('group', { name: 'Propostas do assistente' })
    await expect(proposals.getByText('Alteração no parágrafo 1')).toBeVisible()
    await expect(proposals.getByText('Comentário no parágrafo 1')).toBeVisible()
    await expect(editor.locator('ins.revision')).toHaveCount(0)

    const [first, second] = fake.bodies.slice(-2).map((body) => JSON.parse(body))
    expect(first.tools.map((tool: { function: { name: string } }) => tool.function.name)).toContain(
      'propose_edit',
    )
    expect(first.messages[0].content).toContain('[1] O gato dorme.')
    expect(second.messages.filter((m: { role: string }) => m.role === 'tool')).toHaveLength(2)

    await proposals.getByRole('button', { name: 'Aplicar todas' }).click()
    await expect(editor.locator(`ins.revision[data-author="${TEMPLATE}"]`)).toHaveText('cochila')
    await expect(editor.locator(`del.revision[data-author="${TEMPLATE}"]`)).toHaveText('dorme')
    const card = session.window.locator('.comment-card', { hasText: 'Qual gato?' })
    await expect(card).toContainText(TEMPLATE)
    await expect(proposals.getByText('Aplicada.')).toHaveCount(2)
  })

  test('in an empty document, an insertion writes heading, bold and list as the model revision', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await configure()
    fake.calls = [
      {
        name: 'insert_content',
        arguments: { after: 0, content: '# Guia\n\nTexto com **negrito**.\n\n- Café\n- Chá' },
      },
    ]
    const { field, conversation } = await openConversation()
    await panel().getByLabel(PROPOSE).click()
    await askQuestion(field, 'Escreva um guia.')

    const proposals = conversation.getByRole('group', { name: 'Propostas do assistente' })
    await expect(proposals.getByText('Inserção no início')).toBeVisible()
    await expect(proposals.getByRole('heading', { name: 'Guia' })).toBeVisible()
    await proposals.getByRole('button', { name: 'Aplicar', exact: true }).click()

    await expect(editor.locator('h1')).toHaveText('Guia')
    await expect(editor.locator('h1 ins.revision')).toHaveAttribute('data-author', TEMPLATE)
    await expect(editor.locator('p strong')).toHaveText('negrito')
    await expect(editor.locator('ul li')).toHaveText(['Café', 'Chá'])
    await expect(editor).not.toContainText('#')
    await expect(editor).not.toContainText('**')
  })

  test('Apply and Discard say which proposal they belong to and leave focus in the panel', async () => {
    await write('O gato dorme.')
    await configure()
    fake.calls = [
      { name: 'propose_edit', arguments: { paragraph: 1, find: 'dorme', replace: 'cochila' } },
      { name: 'add_comment', arguments: { paragraph: 1, find: 'gato', comment: 'Qual gato?' } },
    ]
    const { field, conversation } = await openConversation()
    await panel().getByLabel(PROPOSE).click()
    await askQuestion(field, 'Revise o texto.')
    const proposals = conversation.getByRole('group', { name: 'Propostas do assistente' })
    const apply = proposals.getByRole('button', { name: 'Aplicar', exact: true })
    await expect(apply).toHaveCount(2)
    await expect(apply.first()).toHaveAccessibleDescription(/Alteração no parágrafo 1.*dorme.*cochila/u)
    await expect(proposals.getByRole('button', { name: 'Descartar' }).last()).toHaveAccessibleDescription(
      /Comentário no parágrafo 1.*Qual gato\?/u,
    )

    await apply.first().click()
    await expect(proposals.getByText('Aplicada.')).toHaveCount(1)
    await expect(apply).toHaveCount(1)
    await expect(apply).toBeFocused()
    await proposals.getByRole('button', { name: 'Descartar' }).click()
    await expect(proposals.getByText('Descartada.')).toBeVisible()
    await expect(field).toBeFocused()
  })

  test('a model that refuses tools: the reply shows the reason', async () => {
    await write('O gato dorme.')
    await configure()
    fake.withoutTools = true
    const { field, conversation } = await openConversation()
    await panel().getByLabel(PROPOSE).click()
    await askQuestion(field, 'Revise o texto.')
    await expect(conversation.getByText('Sem resposta.')).toBeVisible()
    await expect(
      conversation.getByText('Este modelo não aceita ferramentas.', { exact: false }),
    ).toBeVisible()
  })

  test('with the box unchecked, the conversation request carries no tools', async () => {
    await write('O gato dorme.')
    await configure()
    const { field, conversation } = await openConversation()
    await expect(panel().getByLabel(PROPOSE)).not.toBeChecked()
    await askQuestion(field, 'Qual é o tema?')
    await expect(conversation.getByLabel('Resposta')).toHaveText(REPLY.join(''))
    expect(JSON.parse(fake.bodies.at(-1) ?? '{}')).not.toHaveProperty('tools')
  })

  test('stopping mid-conversation closes the connection and allows sending again', async () => {
    await write('O gato dorme.')
    await configure()
    fake.slow = true
    const { field, conversation } = await openConversation()
    await askQuestion(field, 'Conte devagar.')
    await expect(conversation.getByLabel('Resposta')).toHaveText(/^0 /)

    await panel().getByRole('button', { name: 'Parar' }).click()
    await expect(conversation.getByText('Interrompido.')).toBeVisible()
    await expect.poll(() => fake.closedEarly).toBe(1)
    await expect(conversation.getByRole('button', { name: 'Inserir abaixo' })).toHaveCount(0)
    await field.fill('De novo')
    await expect(panel().getByRole('button', { name: 'Enviar' })).toBeEnabled()
  })
})

/**
 * A fake ACP agent, JSON-RPC written by hand: it thinks, calls a tool, asks permission, and answers
 * only after the permission reply. It logs every message it receives, its pid and its folder. It
 * ignores `session/cancel`, so stopping must end the process. "devagar" in the prompt makes it slow.
 * Given an MCP server in the session, it asks to run Librevia's `add_comment` and, allowed, starts
 * the server as a real agent would and calls it, logging the result.
 */
function fakeAgentSource(logPath: string): string {
  return `#!/usr/bin/env node
import { spawn } from 'node:child_process'
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
let mcp
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
  if (message.method === 'session/new') {
    mcp = message.params.mcpServers[0]
    return send({ id: message.id, result: { sessionId: 's1', configOptions: [MODEL] } })
  }
  if (message.method === 'session/set_config_option') return send({ id: message.id, result: { configOptions: [MODEL] } })
  if (message.method === 'session/prompt' && mcp !== undefined) {
    prompt = message
    const toolCall = { toolCallId: 't2', title: 'mcp__librevia__add_comment', kind: 'other' }
    return send({ id: 'ferramenta', method: 'session/request_permission', params: { sessionId: 's1', toolCall, options: [
      { optionId: 'permitir', name: 'Permitir', kind: 'allow_once' },
      { optionId: 'recusar', name: 'Recusar', kind: 'reject_once' },
    ] } })
  }
  if (message.id === 'ferramenta') {
    if (message.result.outcome.optionId === 'permitir') return callTool()
    return answer(false)
  }
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
function callTool() {
  const env = { ...process.env, ...Object.fromEntries(mcp.env.map(({ name, value }) => [name, value])) }
  const server = spawn(mcp.command, mcp.args, { env, stdio: ['pipe', 'pipe', 'inherit'] })
  const ask = (id, method, params) => server.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\\n')
  let rest = ''
  server.stdout.setEncoding('utf8')
  server.stdout.on('data', (chunk) => {
    const lines = (rest + chunk).split('\\n')
    rest = lines.pop()
    for (const line of lines.filter((text) => text.trim() !== '')) {
      const reply = JSON.parse(line)
      if (reply.id === 1) {
        ask(2, 'tools/call', { name: 'add_comment', arguments: { paragraph: 1, find: 'gato', comment: 'Qual gato?' } })
      } else if (reply.id === 2) {
        log({ toolName: reply.result })
        server.kill()
        answer(false)
      }
    }
  })
  ask(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'agente', version: '1' } })
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
  readonly toolName?: { content: { text: string }[]; isError: boolean }
  readonly pid?: number
  readonly cwd?: string
  readonly id?: string | number
  readonly method?: string
  readonly params?: Record<string, unknown>
  readonly result?: unknown
}

test.describe('assistant through an ACP agent', () => {
  test.skip(process.platform === 'win32', 'the fake agent is a script with a shebang')

  let session: Session
  let directory: string
  let log: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-acp-e2e-'))
    log = join(directory, 'registro.jsonl')
    const agent = join(directory, 'fake-acp-agent.mjs')
    await writeFile(agent, fakeAgentSource(log))
    await chmod(agent, 0o755)
    session = await launch()
    await stubDialogs(session.app, { open: agent, messageBox: 0 })
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  async function readLog(): Promise<AgentLogEntry[]> {
    const text = await readFile(log, 'utf8').catch(() => '')
    return text
      .split('\n')
      .filter((line) => line !== '')
      .map((line) => JSON.parse(line) as AgentLogEntry)
  }

  function panel() {
    return session.window.getByRole('complementary', { name: 'Assistente de IA' })
  }

  async function prepareAgent(text: string): Promise<Locator> {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type(text)
    await menu(session, 'ai-providers')
    const dialog = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    await dialog.getByLabel('Tipo').selectOption({ label: 'Agente ACP (programa instalado)' })
    await dialog.getByLabel('Nome').fill('Agente falso')
    await dialog.getByRole('button', { name: 'Escolher o programa…' }).click()
    await expect(dialog.getByText('Agente falso')).toBeVisible()
    await expect(dialog.getByText(/Fora deste computador · fake-acp-agent\.mjs/)).toBeVisible()
    await dialog.getByRole('button', { name: 'Fechar' }).click()
    await editor.click()
    await session.window.keyboard.press('Control+A')
    await menu(session, 'ai-assistant')
    return editor
  }

  test('the agent replies with text only, the permission is refused and the chosen model becomes the author', async () => {
    const editor = await prepareAgent('O gato dorme.')
    await panel().getByRole('button', { name: 'Listar modelos' }).click()
    const template = panel().getByLabel('Modelo')
    await template.selectOption('agente-esperto')

    await panel().getByRole('button', { name: 'Gerar' }).click()
    const reply = panel().getByLabel('Resposta')
    await expect(reply).toHaveText('Texto do agente.')
    await expect(reply).not.toContainText('pensamento')

    await panel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator('ins.revision[data-author="agente-esperto"]')).toHaveText('Texto do agente.')

    const chatMessages = await readLog()
    const permission = chatMessages.find((message) => message.id === 'permissao')
    expect(permission?.result).toEqual({ outcome: { outcome: 'selected', optionId: 'recusar' } })
    const choice = chatMessages.find((message) => message.method === 'session/set_config_option')
    expect(choice?.params).toMatchObject({ configId: 'model', value: 'agente-esperto' })
    const start = chatMessages.find((message) => message.method === 'initialize')
    expect(start?.params?.['clientCapabilities']).toEqual({
      fs: { readTextFile: false, writeTextFile: false },
      terminal: false,
    })
    const directories = chatMessages.flatMap((message) => (message.cwd === undefined ? [] : [message.cwd]))
    expect(directories.every((cwd) => cwd.includes('librevia-ai-'))).toBe(true)
    await expect.poll(() => directories.some((cwd) => existsSync(cwd))).toBe(false)
  })

  test('with the box checked, the agent uses the Librevia tools through the MCP server and the proposal appears', async () => {
    await prepareAgent('O gato dorme.')
    await panel().getByRole('tab', { name: 'Conversa' }).click()
    await panel().getByLabel('Deixar o assistente propor alterações e comentários').click()
    const field = panel().getByLabel('Mensagem')
    await field.fill('Comente o texto.')
    await field.press('Enter')

    const conversation = panel().getByRole('log')
    await expect(conversation.getByLabel('Resposta')).toHaveText('Texto do agente.')
    const proposals = conversation.getByRole('group', { name: 'Propostas do assistente' })
    await expect(proposals.getByText('Comentário no parágrafo 1')).toBeVisible()
    await expect(session.window.locator('.comment-card')).toHaveCount(0)

    const chatMessages = await readLog()
    const permission = chatMessages.find((message) => message.id === 'ferramenta')
    expect(permission?.result).toEqual({ outcome: { outcome: 'selected', optionId: 'permitir' } })
    const outcome = chatMessages.find((message) => message.toolName !== undefined)?.toolName
    expect(outcome).toMatchObject({ isError: false })
    expect(outcome?.content[0]?.text).toContain('Comment 1 proposed')
  })

  test('stop notifies the agent and ends its process', async () => {
    await prepareAgent('Escreva devagar.')
    await panel().getByRole('button', { name: 'Gerar' }).click()
    await expect(panel().getByLabel('Resposta')).toHaveText(/^0 /)
    await panel().getByRole('button', { name: 'Parar' }).click()
    await expect(panel().getByText('Interrompido.')).toBeVisible()

    await expect
      .poll(async () => (await readLog()).some((message) => message.method === 'session/cancel'))
      .toBe(true)
    const chatMessages = await readLog()
    const pids = chatMessages.flatMap((message) => (message.pid === undefined ? [] : [message.pid]))
    await expect.poll(() => pids.filter(isAlive)).toEqual([])
  })
})

function isAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}
