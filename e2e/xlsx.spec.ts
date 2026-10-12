import { mkdtemp, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { entryOf } from './fixtures.js'

/**
 * The whole `.xlsx` path, which no unit test sees: React → IPC → main → formula translation →
 * binary frame → sidecar → ClosedXML → disk, and back. The spreadsheet is created by the app
 * itself.
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
    // On disk, not just on screen: with the save in progress, closing would open the discard
    // prompt.
    await expect.poll(() => exists(target), { timeout: 30_000 }).toBe(true)

    // Close and reopen: the in-memory model does not prove what is on disk.
    await menu(session, 'close-file')
    await menu(session, 'open')
    await gridReady(session.window)

    await expect(cell(session.window, 2, 0)).toHaveText('37,5')

    // The formula in the bar proves it came back, and in Portuguese.
    await select(session.window, 2, 0)
    await expect(session.window.locator('.formula-bar__input')).toHaveValue('=A1*A2')
  })

  test('as propriedades da planilha vão para o .xlsx e voltam ao reabrir', async () => {
    const target = join(folder, 'orcamento.xlsx')
    session = await launch()
    await stubDialogs(session.app, { save: target, open: target, messageBox: 1 })

    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
    await write(session.window, 0, 0, '3')
    await write(session.window, 1, 0, '=A1*2')

    await menu(session, 'document-properties')
    const dialog = session.window.getByRole('dialog', { name: 'Propriedades da planilha' })
    await expect(dialog.getByRole('row', { name: /Células preenchidas/ })).toContainText('2')
    await expect(dialog.getByRole('row', { name: /Fórmulas/ })).toContainText('1')
    await dialog.getByRole('textbox', { name: 'Título' }).fill('Orçamento 2026')
    await dialog.getByRole('textbox', { name: 'Autor(es)' }).fill('Ana Lima')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialog).toHaveCount(0)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await expect.poll(() => exists(target), { timeout: 30_000 }).toBe(true)

    await menu(session, 'close-file')
    await menu(session, 'open')
    await gridReady(session.window)
    await menu(session, 'document-properties')
    await expect(dialog.getByRole('textbox', { name: 'Título' })).toHaveValue('Orçamento 2026')
    await expect(dialog.getByRole('textbox', { name: 'Autor(es)' })).toHaveValue('Ana Lima')
    // The save stamp: the first revision, and a date where there was none.
    await expect(dialog.getByRole('row', { name: /Revisão/ })).toContainText('1')
    await expect(dialog.getByRole('row', { name: /Modificado em/ })).not.toContainText('—')
  })

  test('o inspetor tira o autor da planilha e o .xlsx guarda o pedido de não gravar nomes', async () => {
    const target = join(folder, 'orcamento.xlsx')
    session = await launch()
    await stubDialogs(session.app, { save: target, messageBox: 1 })

    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
    await write(session.window, 0, 0, '3')
    await menu(session, 'document-properties')
    const properties = session.window.getByRole('dialog', { name: 'Propriedades da planilha' })
    await properties.getByRole('textbox', { name: 'Autor(es)' }).fill('Ana Lima')
    await properties.getByRole('button', { name: 'Aplicar' }).click()
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    await menu(session, 'inspect-document')
    const inspector = session.window.getByRole('dialog', { name: 'Inspecionar documento' })
    await expect(inspector).toContainText('Ana Lima')
    await inspector.getByRole('button', { name: 'Remover' }).click()
    await expect(inspector).toContainText('Removido.')
    await inspector.getByRole('button', { name: 'Fechar' }).click()
    await gridReady(session.window)
    await expect(cell(session.window, 0, 0)).toHaveText('3')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(target, 'xl/workbook.xml')).toContain('filterPrivacy="1"')
    // ClosedXML keeps a new spreadsheet's core properties outside `docProps/`: the dialog reads them.
    await menu(session, 'document-properties')
    await expect(properties.getByRole('textbox', { name: 'Autor(es)' })).toHaveValue('')
    await expect(properties.getByRole('row', { name: /Modificado por/ })).toContainText('—')
  })
})

/**
 * A grid cell, by zero-based coordinates. `revogr-overlay-selection` avoids matching the row
 * header, which uses the same coordinates.
 */
function cell(window: Page, row: number, column: number) {
  return window
    .locator(`revogr-overlay-selection revogr-data [data-rgrow="${row}"][data-rgcol="${column}"]`)
    .first()
}

/** The grid only draws cells after mounting; clicking earlier selects nothing. */
async function gridReady(window: Page): Promise<void> {
  await expect(window.locator('revo-grid')).toBeVisible()
  await expect(cell(window, 0, 0)).toBeVisible()
}

async function select(window: Page, row: number, column: number): Promise<void> {
  // The click repeats until the selection moves: the grid redraws after opening.
  await expect(async () => {
    await cell(window, row, column).click()
    await expect(window.locator('.formula-bar__ref')).toHaveText(reference(row, column), {
      timeout: 1000,
    })
  }).toPass({ timeout: 15_000 })
}

/** Selects the cell and writes into it through the formula bar. */
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
