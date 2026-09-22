import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTextBox } from './fixtures.js'

let session: Session | undefined
let folder: string
let file: string

test.beforeEach(async () => {
  folder = await mkdtemp(join(tmpdir(), 'librevia-external-'))
  file = join(folder, 'Document with spaces.DOCX')
  await writeFile(file, await docxWithTextBox())
})

test.afterEach(async () => {
  if (session !== undefined) {
    await session.app.close()
    await rm(session.userData, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
  }
  session = undefined
  await rm(folder, { recursive: true, force: true })
})

test('opens a DOCX passed at startup after the renderer is ready', async () => {
  session = await launch({ file })
  await expect(session.window.locator('.ProseMirror')).toContainText('Título na caixa')
  await expect(session.window.locator('.statusbar')).toContainText(file)
})

test('opens an Explorer request in the existing window and respects cancel', async () => {
  session = await launch()
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').fill('Unsaved work')
  await stubDialogs(session.app, { messageBox: 2 })

  const request = async (): Promise<void> => {
    await session!.app.evaluate(({ app }, path) => {
      const args = app.isPackaged ? [process.execPath, path] : [process.execPath, 'app', path]
      app.emit('second-instance', {}, args, process.cwd())
    }, file)
  }

  await request()
  await expect(session.window.locator('.ProseMirror')).toHaveText('Unsaved work')
  await stubDialogs(session.app, { messageBox: 1 })
  await request()
  await expect(session.window.locator('.ProseMirror')).toContainText('Título na caixa')
  expect(session.app.windows()).toHaveLength(1)
})
