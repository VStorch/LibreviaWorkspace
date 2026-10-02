import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithFootnote, entryOf } from './fixtures.js'

/**
 * Notas de rodapé e de fim (M11, fase 1): lidas, numeradas e preservadas.
 *
 * A referência é um nó com o corpo da nota dentro. Editar o parágrafo que a leva
 * não a perde mais — e por isso o documento com nota deixou de abrir travado.
 */
test.describe('notas de rodapé', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-notas-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('editar o parágrafo da referência, salvar e reabrir mantém a nota', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithFootnote())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')

    await editor.getByText('Segundo parágrafo').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Revisto: ')
    await expect(editor).toContainText('Revisto: Segundo parágrafo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // O parágrafo foi reescrito com a referência; a nota, que ninguém tocou,
    // continua no arquivo, com os separadores.
    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('Revisto: Segundo parágrafo')
    expect(corpo).toMatch(/<w:footnoteReference w:id="1" ?\/>/)
    const notas = await entryOf(origem, 'word/footnotes.xml')
    expect(notas).toContain('Fonte: ata anterior.')
    expect(notas).toContain('w:separator')

    await menu(session, 'open')
    await expect(editor).toContainText('Revisto: Segundo parágrafo')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('title', /Fonte: ata anterior\./)
    await expect(session.window.locator('.banner--notice')).toBeHidden()
  })
})
