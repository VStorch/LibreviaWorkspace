import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTrackedChange } from './fixtures.js'
import { cell, gridReady, write } from './sheet-grid.js'

const CLIENTE = 'Cliente de teste'
const OPCAO = 'Permitir que programas externos leiam o documento aberto e proponham alterações (MCP)'

interface ClienteMcp {
  readonly processo: ChildProcessWithoutNullStreams
  readonly saiu: Promise<number | null>
  pedir(method: string, params?: unknown): Promise<Record<string, unknown>>
}

/** The bridge as a client would start it, from the configuration the app shows, in the same data folder. */
function abrirPonte(comando: string, args: readonly string[], userData: string): ClienteMcp {
  const processo = spawn(comando, [...args, `--user-data-dir=${userData}`], {
    env: { ...process.env, NODE_ENV: 'production' },
  })
  const saiu = new Promise<number | null>((resolve) => processo.once('exit', (code) => resolve(code)))
  const esperando = new Map<number, (mensagem: Record<string, unknown>) => void>()
  let resto = ''
  let proximo = 1
  processo.stdout.setEncoding('utf8')
  processo.stdout.on('data', (pedaco: string) => {
    const linhas = (resto + pedaco).split('\n')
    resto = linhas.pop() ?? ''
    for (const linha of linhas.filter((texto) => texto.trim() !== '')) {
      const mensagem = JSON.parse(linha) as Record<string, unknown>
      esperando.get(mensagem['id'] as number)?.(mensagem)
    }
  })
  return {
    processo,
    saiu,
    pedir: (method, params) => {
      const id = proximo++
      processo.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`)
      return new Promise((resolve) => esperando.set(id, resolve))
    },
  }
}

const textoDe = (resposta: Record<string, unknown>): string =>
  (resposta['result'] as { content: { text: string }[] }).content[0]!.text

test.describe('servidor MCP local', () => {
  let session: Session
  let pasta: string
  const pontes: ClienteMcp[] = []

  test.beforeEach(async () => {
    session = await launch()
    pasta = await mkdtemp(join(tmpdir(), 'librevia-mcp-'))
  })

  test.afterEach(async () => {
    for (const ponte of pontes.splice(0)) ponte.processo.kill()
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  async function definirPreferencia(ligar: boolean): Promise<{ command: string; args: string[] } | null> {
    await menu(session, 'ai-providers')
    const dialogo = session.window.getByRole('dialog', { name: 'Assistentes de IA' })
    const caixa = dialogo.getByLabel(OPCAO)
    await caixa.click()
    await expect(caixa).toBeChecked({ checked: ligar })
    let configuracao: { command: string; args: string[] } | null = null
    if (ligar) {
      await expect(dialogo.getByTestId('mcp-setup-command').locator('pre')).toContainText(
        'claude mcp add --scope user librevia --',
      )
      const texto = dialogo.getByTestId('mcp-setup-json').locator('pre')
      await expect(texto).toContainText('--mcp-bridge')
      const json = JSON.parse(await texto.innerText()) as {
        mcpServers: { librevia: { command: string; args: string[] } }
      }
      configuracao = json.mcpServers.librevia
    }
    await dialogo.getByRole('button', { name: 'Fechar' }).click()
    return configuracao
  }

  test('o cliente lê o documento e a proposta entra como revisão com o autor (MCP); desligado, recusa', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Ata da reunião de terça.')

    const configuracao = await definirPreferencia(true)
    expect(configuracao).not.toBeNull()

    const ponte = abrirPonte(configuracao!.command, configuracao!.args, session.userData)
    pontes.push(ponte)
    const inicio = await ponte.pedir('initialize', {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: { name: CLIENTE, version: '1' },
    })
    expect(inicio['result']).toMatchObject({ serverInfo: { name: 'librevia' } })

    const lista = await ponte.pedir('tools/list')
    const nomes = (lista['result'] as { tools: { name: string }[] }).tools.map((tool) => tool.name)
    expect(nomes).toEqual([
      'read_paragraphs',
      'search_document',
      'propose_edit',
      'add_comment',
      'list_sheets',
      'read_cells',
      'find_cells',
      'write_cells',
    ])

    const leitura = await ponte.pedir('tools/call', {
      name: 'read_paragraphs',
      arguments: { from: 1, count: 2 },
    })
    expect(textoDe(leitura)).toContain('[1] Ata da reunião de terça.')

    const proposta = await ponte.pedir('tools/call', {
      name: 'propose_edit',
      arguments: { paragraph: 1, find: 'terça', replace: 'quarta' },
    })
    expect(textoDe(proposta)).toContain('tracked revision')
    await expect(editor.locator(`ins.revision[data-author="${CLIENTE} (MCP)"]`)).toHaveText('quarta')
    await expect(editor.locator(`del.revision[data-author="${CLIENTE} (MCP)"]`)).toHaveText('terça')

    const indicador = session.window.getByTestId('mcp-status')
    await expect(indicador).toContainText(CLIENTE)

    await definirPreferencia(false)
    expect(await ponte.saiu).toBe(0)
    await expect(indicador).toHaveCount(0)

    const recusada = abrirPonte(configuracao!.command, configuracao!.args, session.userData)
    pontes.push(recusada)
    expect(await recusada.saiu).toBe(1)
  })

  test('Desconectar na barra de status derruba o cliente e deixa o servidor ligado', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Ata da reunião de terça.')
    const configuracao = await definirPreferencia(true)

    const ponte = abrirPonte(configuracao!.command, configuracao!.args, session.userData)
    pontes.push(ponte)
    await ponte.pedir('initialize', {
      protocolVersion: '2025-06-18',
      clientInfo: { name: CLIENTE, version: '1' },
    })
    const indicador = session.window.getByTestId('mcp-status')
    await expect(indicador).toContainText(CLIENTE)

    await indicador.getByRole('button', { name: 'Desconectar' }).click()
    expect(await ponte.saiu).toBe(0)
    await expect(indicador).toHaveCount(0)

    const outra = abrirPonte(configuracao!.command, configuracao!.args, session.userData)
    pontes.push(outra)
    const inicio = await outra.pedir('initialize', {
      protocolVersion: '2025-06-18',
      clientInfo: { name: CLIENTE, version: '1' },
    })
    expect(inicio['result']).toMatchObject({ serverInfo: { name: 'librevia' } })
  })

  test('na planilha, o cliente lê as células e a escrita espera a pessoa aplicar ou descartar', async () => {
    const { window } = session
    await menu(session, 'new-spreadsheet')
    await gridReady(window)
    await write(window, 0, 0, 'Produto')
    await write(window, 1, 0, 'Café')
    await write(window, 1, 1, '=2*5')
    const configuracao = await definirPreferencia(true)
    const ponte = abrirPonte(configuracao!.command, configuracao!.args, session.userData)
    pontes.push(ponte)
    await ponte.pedir('initialize', {
      protocolVersion: '2025-06-18',
      clientInfo: { name: CLIENTE, version: '1' },
    })

    const chamar = async (name: string, args: unknown): Promise<string> =>
      textoDe(await ponte.pedir('tools/call', { name, arguments: args }))
    expect(await chamar('read_cells', { range: 'A1:B2' })).toBe(
      ['Sheet: Planilha1', 'A1: Produto', 'A2: Café', 'B2: =2*5 → 10'].join('\n'),
    )
    expect(await chamar('read_paragraphs', { from: 1, count: 1 })).toContain('No text document is open')

    expect(await chamar('write_cells', { start: 'B1', rows: [['Total'], ['=B3*2'], ['7']] })).toContain(
      'waiting',
    )
    const faixa = window.getByRole('region', { name: 'Proposta de um programa externo' })
    await expect(faixa).toContainText(`${CLIENTE} (MCP) propõe escrever em Planilha1!B1:B3`)
    await expect(faixa).toContainText('Substitui 1 célula preenchida.')
    await expect(faixa.getByRole('table')).toContainText('=B3*2')
    await expect(cell(window, 0, 1)).toHaveText('')

    await faixa.getByRole('button', { name: 'Aplicar' }).click()
    await expect(faixa).toHaveCount(0)
    await expect(cell(window, 0, 1)).toHaveText('Total')
    await expect(cell(window, 1, 1)).toHaveText('14')
    await window.keyboard.press('Control+z')
    await expect(cell(window, 1, 1)).toHaveText('10')
    await expect(cell(window, 0, 1)).toHaveText('')

    await chamar('write_cells', { start: 'C1', rows: [['x']] })
    await faixa.getByRole('button', { name: 'Descartar' }).click()
    await expect(faixa).toHaveCount(0)
    await expect(cell(window, 0, 2)).toHaveText('')
  })
})
