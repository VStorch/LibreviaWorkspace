import { expect, test } from '@playwright/test'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderGrid, entryOf } from './fixtures.js'

/**
 * A document born in the editor, saved as `.docx` over the sidecar's minimal package, which every
 * save starts from. Along the whole path: dialog, published sidecar, disk.
 */
test.describe('documento novo em .docx', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-docx-novo-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('salva como .docx, grava de novo no mesmo arquivo e reabre com título e texto', async () => {
    const destino = join(folder, 'relatorio.docx')
    const editor = session.window.locator('.ProseMirror')

    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('Relatório anual')
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Título 1' })

    await editor.locator('h1').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Primeiro parágrafo.')

    await stubDialogs(session.app, { save: destino, open: destino })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // The heading points to a style the package defines, not a loose name.
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toContain('<w:pStyle w:val="Heading1"')
    expect(corpo).toContain('Primeiro parágrafo.')
    expect(await entryOf(destino, 'word/styles.xml')).toContain('w:styleId="Heading1"')

    // The second save goes to the same file without asking: the path was authorized, and it starts
    // from the same minimal package as the first.
    await session.window.keyboard.type(' Segunda gravação.')
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(destino, 'word/document.xml')).toContain('Segunda gravação.')

    await menu(session, 'open')
    await expect(editor.locator('h1')).toHaveText('Relatório anual')
    await expect(editor.locator('p').first()).toHaveText('Primeiro parágrafo. Segunda gravação.')
  })

  test('o .sdoc que veio de um .docx com cabeçalho volta a .docx avisando da faixa', async () => {
    // The `.sdoc` keeps the band with the source `.docx` relationships, which the minimal package
    // lacks: the band is lost, and the file must go out with the loss stated.
    const origem = join(folder, 'grade.docx')
    const rascunho = join(folder, 'grade.sdoc')
    const destino = join(folder, 'volta.docx')
    const editor = session.window.locator('.ProseMirror')
    await writeFile(origem, await docxWithHeaderGrid())

    await stubDialogs(session.app, { open: origem, save: rascunho, messageBox: 1 })
    await menu(session, 'open')
    await expect(editor).toContainText('Primeira linha do corpo.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Reopening from disk is what drops the original: the app keeps the `.docx` bytes on open, and
    // whoever opens the `.sdoc` does not have them.
    await stubDialogs(session.app, { open: rascunho })
    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()
    await menu(session, 'open')
    await expect(editor).toContainText('Primeira linha do corpo.')

    await stubDialogs(session.app, { save: destino })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // A loss notice, not an error banner: `role="alert"` belongs to ErrorBanner.
    const aviso = session.window.locator('.banner--notice')
    await expect(aviso).toContainText('cabeçalho e rodapé do arquivo .docx de origem')
    await expect(session.window.getByRole('alert')).toHaveCount(0)

    // And the file really exists, with the body inside.
    expect(await entryOf(destino, 'word/document.xml')).toContain('Primeira linha do corpo.')
  })
})
