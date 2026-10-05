import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/** The file cycle, in the assembled app, sandboxed preload and all. */
test.describe('ciclo de arquivo', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-basics-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('abre na tela inicial', async () => {
    await expect(session.window.locator('.home')).toBeVisible()
  })

  test('escrever, salvar, fechar e reabrir devolve o texto', async () => {
    const target = join(folder, 'ata.sdoc')
    await stubDialogs(session.app, { save: target, open: target, messageBox: 1 })

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Presentes: Ana, Bruno e Carla.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()

    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Presentes: Ana, Bruno e Carla.')
  })

  test('o formato do arquivo decide o editor', async () => {
    // `.ssheet` must open in the grid, not in the text editor showing JSON.
    const target = join(folder, 'contas.ssheet')
    await stubDialogs(session.app, { save: target, open: target, messageBox: 1 })

    await menu(session, 'new-spreadsheet')
    await menu(session, 'save-as')
    // Waiting for the save to finish is not fussiness: `menu` only fires the command, and without
    // this `open` raced `save-as` and read a file that did not exist yet; the test passed because
    // the new spreadsheet grid was still on screen, not because the file had opened.
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()

    await menu(session, 'open')

    await expect(session.window.locator('revo-grid')).toBeVisible()
    await expect(session.window.locator('.ProseMirror')).toBeHidden()
  })
})
