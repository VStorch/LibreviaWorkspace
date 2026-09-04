import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderTextBox, entryOf } from './fixtures.js'

/**
 * O cabeçalho que é um grupo de formas.
 *
 * A maioria do corpus é assim: o título não está num parágrafo do cabeçalho,
 * está dentro de uma caixa ancorada, ao lado do logotipo e do campo do número
 * da página. Quem via um título errado ali não tinha onde clicar.
 *
 * A caixa volta inteira, e não peça por peça: digitar dentro dela abre e fecha
 * parágrafos. Já o campo `PAGE` não volta de jeito nenhum — ele é calculado a
 * cada abertura, e escrever o marcador de volta o trocaria por um número fixo.
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

    // O grupo segue inteiro, com o campo que este escritor não sabe gerar.
    expect(cabecalho).toContain('PAGE')
    expect(cabecalho).toContain('fldChar')
  })

  test('a caixa do número da página não recebe o cursor', async () => {
    // O que se vê nela é o número desta folha; o que está no arquivo é um
    // campo. Devolvê-lo como texto faria o cabeçalho dizer "1" em todas as
    // folhas, e só se notaria na segunda.
    const numero = session.window.locator('.paper-float--text').filter({ hasText: /^1$/ })
    await expect(numero).toHaveCount(1)
    await expect(numero).not.toHaveClass(/paper-float--edit/)
  })
})
