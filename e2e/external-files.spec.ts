import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
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

test('the status bar shows the path in order and cuts a long one at the start', async () => {
  const deep = join(folder, 'uma pasta com um nome bem comprido'.repeat(4), 'outra pasta')
  await mkdir(deep, { recursive: true })
  const long = join(deep, 'Ata final.docx')
  await writeFile(long, await docxWithTextBox())

  session = await launch({ file })
  await expect(session.window.locator('.statusbar')).toContainText(file)
  // Bidi under `rtl` drew the leading slash after the name: `tmp/…/name.docx/`.
  expect(await drawnInOrder(session.window)).toBe(true)

  await stubDialogs(session.app, { open: long })
  await menu(session, 'open')
  await expect(session.window.locator('.statusbar')).toContainText(long)
  expect(await nameStaysVisible(session.window)).toBe(true)
})

/** The first character is drawn left of the last. */
async function drawnInOrder(page: Page): Promise<boolean> {
  return page.locator('.statusbar__file bdi').evaluate((element) => {
    const text = element.firstChild!
    const range = document.createRange()
    const charAt = (index: number): DOMRect => {
      range.setStart(text, index)
      range.setEnd(text, index + 1)
      return range.getBoundingClientRect()
    }
    return charAt(0).left < charAt(text.textContent!.length - 1).left
  })
}

/** The cut falls on the start of the path: the end, with the name, stays inside the box. */
async function nameStaysVisible(page: Page): Promise<boolean> {
  return page.locator('.statusbar__file').evaluate((box) => {
    const text = box.querySelector('bdi')!.firstChild!
    const range = document.createRange()
    range.setStart(text, text.textContent!.length - 1)
    range.setEnd(text, text.textContent!.length)
    const last = range.getBoundingClientRect()
    const bounds = box.getBoundingClientRect()
    return box.scrollWidth > box.clientWidth && last.right <= bounds.right + 1 && last.left >= bounds.left
  })
}
