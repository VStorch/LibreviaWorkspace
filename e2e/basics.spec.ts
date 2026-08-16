import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'

/** O ciclo de arquivo, no aplicativo montado — com preload sandboxed e tudo. */
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
    // `.ssheet` precisa abrir na grade, e não no editor de texto mostrando JSON.
    const target = join(folder, 'contas.ssheet')
    await stubDialogs(session.app, { save: target, open: target, messageBox: 1 })

    await menu(session, 'new-spreadsheet')
    await menu(session, 'save-as')
    await menu(session, 'close-file')
    await menu(session, 'open')

    await expect(session.window.locator('revo-grid')).toBeVisible()
    await expect(session.window.locator('.ProseMirror')).toBeHidden()
  })
})
