import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, selectionSettled, stubDialogs, type Session } from './app.js'
import { entryOf, odtWithParagraphs } from './fixtures.js'

/**
 * An `.odt` opens as a file of its own: Ctrl+S writes into it, and only the edited paragraph is
 * written again. The other goes back as the exact text it had in `content.xml`. Any other document
 * saves as `.odt` too, whole, into a new package.
 */
test.describe('ODT', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-odt-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('editing a paragraph and pressing Ctrl+S saves into the same .odt', async () => {
    const file = join(folder, 'ata.odt')
    await writeFile(file, await odtWithParagraphs('Primeiro parágrafo intacto.', 'Segundo parágrafo.'))
    await stubDialogs(session.app, { open: file, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(session.window).toHaveTitle(/ata\.odt/)
    await editor.getByText('Segundo parágrafo.').click({ position: { x: 4, y: 4 } })
    await selectionSettled(session)
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Revisto: ')
    await expect(editor).toContainText('Revisto: Segundo parágrafo.')

    // The command Ctrl+S fires: a synthetic key does not reach the native menu's accelerator.
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const content = await entryOf(file, 'content.xml')
    expect(content).toContain('<text:p text:style-name="Standard">Primeiro parágrafo intacto.</text:p>')
    expect(content).toContain('Revisto: Segundo parágrafo.')

    await menu(session, 'open')
    await expect(editor).toContainText('Revisto: Segundo parágrafo.')
  })

  test('a new document saved as .odt opens again with its text', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.pages__column .ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Documento novo em ODT.')

    const file = join(folder, 'novo.odt')
    await stubDialogs(session.app, { save: file })
    await menu(session, 'save-as')
    await expect(session.window).toHaveTitle(/^novo\.odt/)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(file, 'content.xml')).toContain('Documento novo em ODT.')

    await menu(session, 'new-document')
    await expect(editor).not.toContainText('Documento novo em ODT.')
    await stubDialogs(session.app, { open: file, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window).toHaveTitle(/^novo\.odt/)
    await expect(editor).toContainText('Documento novo em ODT.')
  })
})
