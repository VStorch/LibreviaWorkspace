import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTrackedChange } from './fixtures.js'
import { cell, gridReady, write } from './sheet-grid.js'

const CLIENT = 'Cliente de teste'
const OPTION = 'Permitir que programas externos leiam o documento aberto e proponham alterações (MCP)'

interface McpClient {
  readonly process: ChildProcessWithoutNullStreams
  readonly left: Promise<number | null>
  ask(method: string, params?: unknown): Promise<Record<string, unknown>>
}

/** The bridge as a client would start it, from the configuration the app shows, in the same data folder. */
function openBridge(command: string, args: readonly string[], userData: string): McpClient {
  const child = spawn(command, [...args, `--user-data-dir=${userData}`], {
    env: { ...process.env, NODE_ENV: 'production' },
  })
  const left = new Promise<number | null>((resolve) => child.once('exit', (code) => resolve(code)))
  const waiting = new Map<number, (message: Record<string, unknown>) => void>()
  let rest = ''
  let next = 1
  child.stdout.setEncoding('utf8')
  child.stdout.on('data', (chunk: string) => {
    const lines = (rest + chunk).split('\n')
    rest = lines.pop() ?? ''
    for (const line of lines.filter((text) => text.trim() !== '')) {
      const message = JSON.parse(line) as Record<string, unknown>
      waiting.get(message['id'] as number)?.(message)
    }
  })
  return {
    process: child,
    left,
    ask: (method, params) => {
      const id = next++
      child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`)
      return new Promise((resolve) => waiting.set(id, resolve))
    },
  }
}

const textOf = (reply: Record<string, unknown>): string =>
  (reply['result'] as { content: { text: string }[] }).content[0]!.text

test.describe('local MCP server', () => {
  let session: Session
  let directory: string
  const bridges: McpClient[] = []

  test.beforeEach(async () => {
    session = await launch()
    directory = await mkdtemp(join(tmpdir(), 'librevia-mcp-'))
  })

  test.afterEach(async () => {
    for (const bridge of bridges.splice(0)) bridge.process.kill()
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  async function setServerPreference(enable: boolean): Promise<{ command: string; args: string[] } | null> {
    await menu(session, 'ai-providers')
    const dialog = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    const box = dialog.getByLabel(OPTION)
    await box.click()
    await expect(box).toBeChecked({ checked: enable })
    let configuration: { command: string; args: string[] } | null = null
    if (enable) {
      await expect(dialog.getByTestId('mcp-setup-command').locator('pre')).toContainText(
        'claude mcp add --scope user librevia --',
      )
      const text = dialog.getByTestId('mcp-setup-json').locator('pre')
      await expect(text).toContainText('--mcp-bridge')
      const json = JSON.parse(await text.innerText()) as {
        mcpServers: { librevia: { command: string; args: string[] } }
      }
      configuration = json.mcpServers.librevia
    }
    await dialog.getByRole('button', { name: 'Fechar' }).click()
    return configuration
  }

  test('the client reads the document and the proposal enters as a revision with the (MCP) author; when off, it refuses', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Ata da reunião de terça.')

    const configuration = await setServerPreference(true)
    expect(configuration).not.toBeNull()

    const bridge = openBridge(configuration!.command, configuration!.args, session.userData)
    bridges.push(bridge)
    const start = await bridge.ask('initialize', {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: { name: CLIENT, version: '1' },
    })
    expect(start['result']).toMatchObject({ serverInfo: { name: 'librevia' } })

    const list = await bridge.ask('tools/list')
    const names = (list['result'] as { tools: { name: string }[] }).tools.map((tool) => tool.name)
    expect(names).toEqual([
      'read_paragraphs',
      'search_document',
      'propose_edit',
      'add_comment',
      'list_sheets',
      'read_cells',
      'find_cells',
      'write_cells',
    ])

    const reading = await bridge.ask('tools/call', {
      name: 'read_paragraphs',
      arguments: { from: 1, count: 2 },
    })
    expect(textOf(reading)).toContain('[1] Ata da reunião de terça.')

    const proposal = await bridge.ask('tools/call', {
      name: 'propose_edit',
      arguments: { paragraph: 1, find: 'terça', replace: 'quarta' },
    })
    expect(textOf(proposal)).toContain('tracked revision')
    await expect(editor.locator(`ins.revision[data-author="${CLIENT} (MCP)"]`)).toHaveText('quarta')
    await expect(editor.locator(`del.revision[data-author="${CLIENT} (MCP)"]`)).toHaveText('terça')

    const indicator = session.window.getByTestId('mcp-status')
    await expect(indicator).toContainText(CLIENT)

    await setServerPreference(false)
    expect(await bridge.left).toBe(0)
    await expect(indicator).toHaveCount(0)

    const refused = openBridge(configuration!.command, configuration!.args, session.userData)
    bridges.push(refused)
    expect(await refused.left).toBe(1)
  })

  test('Disconnect in the status bar drops the client and leaves the server on', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Ata da reunião de terça.')
    const configuration = await setServerPreference(true)

    const bridge = openBridge(configuration!.command, configuration!.args, session.userData)
    bridges.push(bridge)
    await bridge.ask('initialize', {
      protocolVersion: '2025-06-18',
      clientInfo: { name: CLIENT, version: '1' },
    })
    const indicator = session.window.getByTestId('mcp-status')
    await expect(indicator).toContainText(CLIENT)

    await indicator.getByRole('button', { name: 'Desconectar' }).click()
    expect(await bridge.left).toBe(0)
    await expect(indicator).toHaveCount(0)

    const other = openBridge(configuration!.command, configuration!.args, session.userData)
    bridges.push(other)
    const start = await other.ask('initialize', {
      protocolVersion: '2025-06-18',
      clientInfo: { name: CLIENT, version: '1' },
    })
    expect(start['result']).toMatchObject({ serverInfo: { name: 'librevia' } })
  })

  test('in the spreadsheet, the client reads the cells and writing waits for the person to apply or discard', async () => {
    const { window } = session
    await menu(session, 'new-spreadsheet')
    await gridReady(window)
    await write(window, 0, 0, 'Produto')
    await write(window, 1, 0, 'Café')
    await write(window, 1, 1, '=2*5')
    const configuration = await setServerPreference(true)
    const bridge = openBridge(configuration!.command, configuration!.args, session.userData)
    bridges.push(bridge)
    await bridge.ask('initialize', {
      protocolVersion: '2025-06-18',
      clientInfo: { name: CLIENT, version: '1' },
    })

    const callTool = async (name: string, args: unknown): Promise<string> =>
      textOf(await bridge.ask('tools/call', { name, arguments: args }))
    expect(await callTool('read_cells', { range: 'A1:B2' })).toBe(
      ['Sheet: Planilha1', 'A1: Produto', 'A2: Café', 'B2: =2*5 → 10'].join('\n'),
    )
    expect(await callTool('read_paragraphs', { from: 1, count: 1 })).toContain('No text document is open')

    expect(await callTool('write_cells', { start: 'B1', rows: [['Total'], ['=B3*2'], ['7']] })).toContain(
      'waiting',
    )
    const band = window.getByRole('region', { name: 'Proposta de um programa externo' })
    await expect(band).toContainText(`${CLIENT} (MCP) propõe escrever em Planilha1!B1:B3`)
    await expect(band).toContainText('Substitui 1 célula preenchida.')
    await expect(band.getByRole('table')).toContainText('=B3*2')
    await expect(cell(window, 0, 1)).toHaveText('')

    await band.getByRole('button', { name: 'Aplicar' }).click()
    await expect(band).toHaveCount(0)
    await expect(cell(window, 0, 1)).toHaveText('Total')
    await expect(cell(window, 1, 1)).toHaveText('14')
    await window.keyboard.press('Control+z')
    await expect(cell(window, 1, 1)).toHaveText('10')
    await expect(cell(window, 0, 1)).toHaveText('')

    await callTool('write_cells', { start: 'C1', rows: [['x']] })
    await band.getByRole('button', { name: 'Descartar' }).click()
    await expect(band).toHaveCount(0)
    await expect(cell(window, 0, 2)).toHaveText('')
  })
})
