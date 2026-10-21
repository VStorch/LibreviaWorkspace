import { mkdtemp, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { entryOf } from './fixtures.js'
import { cell, gridReady, select, write } from './sheet-grid.js'

/**
 * The whole `.xlsx` path, which no unit test sees: React → IPC → main → formula translation →
 * binary frame → sidecar → ClosedXML → disk, and back. The spreadsheet is created by the app
 * itself.
 */
test.describe('spreadsheet in .xlsx', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-xlsx-'))
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('saves, reopens and the formulas stay formulas', async () => {
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

  test('the spreadsheet properties go to the .xlsx and come back on reopening', async () => {
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

  test('the inspector removes the spreadsheet author and the .xlsx keeps the request not to store names', async () => {
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

async function exists(path: string): Promise<boolean> {
  return stat(path).then(
    () => true,
    () => false,
  )
}
