import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderTextBox, entryOf } from './fixtures.js'

/**
 * A header that is a shape group, like most of the corpus: the title in an anchored box. The box
 * comes back whole, because typing in it opens and closes paragraphs; the `PAGE` field does not, or
 * it would become a fixed number.
 */
test.describe('caixa de cabeçalho editável', () => {
  let session: Session
  let pasta: string
  let destino: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-faixa-'))
    destino = join(pasta, 'salva.docx')
    session = await launch()

    const origem = join(pasta, 'grupo.docx')
    await writeFile(origem, await docxWithHeaderTextBox())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toBeVisible()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('o título da caixa é digitável e volta para o cabeçalho do arquivo', async () => {
    const caixa = session.window
      .locator('.paper-float--text')
      .filter({ hasText: 'EVIDÊNCIAS DO ROTEIRO' })
      .first()

    await caixa.click()
    await session.window.keyboard.press('ControlOrMeta+a')
    await session.window.keyboard.type('EVIDÊNCIAS DE HOMOLOGAÇÃO')
    await session.window.locator('.ProseMirror').click()

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const cabecalho = await entryOf(destino, 'word/header1.xml')
    expect(cabecalho).toContain('EVIDÊNCIAS DE HOMOLOGAÇÃO')
    expect(cabecalho).not.toContain('EVIDÊNCIAS DO ROTEIRO')

    // The group stays whole, with the field this writer cannot generate.
    expect(cabecalho).toContain('PAGE')
    expect(cabecalho).toContain('fldChar')
  })

  test('a caixa do número da página não recebe o cursor', async () => {
    // What shows in it is this sheet's number; what is in the file is a field. Returning it as text
    // would make the header say "1" on every sheet, and it would only be noticed on the second.
    const numero = session.window.locator('.paper-float--text').filter({ hasText: /^1$/ })
    await expect(numero).toHaveCount(1)
    await expect(numero).not.toHaveClass(/paper-float--edit/)
  })
})
