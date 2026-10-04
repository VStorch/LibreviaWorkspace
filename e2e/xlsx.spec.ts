import { mkdtemp, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/**
 * O caminho inteiro do `.xlsx`, que nenhum teste de unidade vê: React → IPC → main →
 * tradução de fórmula → quadro binário → sidecar → ClosedXML → disco, e a volta. A
 * planilha é criada pelo próprio aplicativo.
 */
test.describe('planilha em .xlsx', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-xlsx-'))
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('salva, reabre e as fórmulas continuam fórmulas', async () => {
    const target = join(folder, 'vendas.xlsx')
    session = await launch()
    await stubDialogs(session.app, { save: target, open: target, messageBox: 1 })

    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)

    await write(session.window, 0, 0, '3')
    await write(session.window, 1, 0, '12,5')
    await write(session.window, 2, 0, '=A1*A2')

    await expect(cell(session.window, 2, 0)).toHaveText('37,5')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    // No disco, e não só na tela: com a gravação em curso, fechar abriria o aviso de descarte.
    await expect.poll(() => exists(target), { timeout: 30_000 }).toBe(true)

    // Fechar e reabrir: o modelo da memória não prova o que está no disco.
    await menu(session, 'close-file')
    await menu(session, 'open')
    await gridReady(session.window)

    await expect(cell(session.window, 2, 0)).toHaveText('37,5')

    // A fórmula na barra prova que ela voltou, e em português.
    await select(session.window, 2, 0)
    await expect(session.window.locator('.formula-bar__input')).toHaveValue('=A1*A2')
  })
})

/**
 * Uma célula da grade, pelas coordenadas base zero. O `revogr-overlay-selection`
 * evita casar com o cabeçalho de linha, que usa as mesmas coordenadas.
 */
function cell(window: Page, row: number, column: number) {
  return window
    .locator(`revogr-overlay-selection revogr-data [data-rgrow="${row}"][data-rgcol="${column}"]`)
    .first()
}

/** A grade só desenha as células depois de montar; clicar antes não seleciona. */
async function gridReady(window: Page): Promise<void> {
  await expect(window.locator('revo-grid')).toBeVisible()
  await expect(cell(window, 0, 0)).toBeVisible()
}

async function select(window: Page, row: number, column: number): Promise<void> {
  // O clique repete até a seleção mover: a grade se redesenha depois de abrir.
  await expect(async () => {
    await cell(window, row, column).click()
    await expect(window.locator('.formula-bar__ref')).toHaveText(reference(row, column), {
      timeout: 1000,
    })
  }).toPass({ timeout: 15_000 })
}

/** Seleciona a célula e escreve nela pela barra de fórmulas. */
async function write(window: Page, row: number, column: number, text: string): Promise<void> {
  await select(window, row, column)
  const input = window.locator('.formula-bar__input')
  await input.fill(text)
  await input.press('Enter')
}

async function exists(path: string): Promise<boolean> {
  return stat(path).then(
    () => true,
    () => false,
  )
}

function reference(row: number, column: number): string {
  return `${String.fromCharCode(65 + column)}${row + 1}`
}
