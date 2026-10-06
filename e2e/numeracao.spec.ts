import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithPageNumbering, entryOf } from './fixtures.js'
import { hasPdftotext, pdfText } from './external-tools.js'

/**
 * Page numbering: each sheet's number with the start and format the document asks for, the field
 * the user inserted, and the band switches.
 */
test.describe('numeração de página', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-numeracao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  const rodapes = (session: Session) =>
    session.window
      .locator('.band--footer')
      .evaluateAll((faixas) => faixas.map((faixa) => (faixa.textContent ?? '').trim()))

  async function abrir(): Promise<string> {
    const origem = join(pasta, 'entrada.docx')
    await writeFile(origem, await docxWithPageNumbering())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(2)
    return origem
  }

  test('cada folha mostra o número no formato e a partir do início do documento', async () => {
    await abrir()
    await expect.poll(() => rodapes(session)).toEqual(['Página iii', 'Página iv'])
  })

  test('o PDF numera as folhas como a tela', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext não instalado')
    await abrir()
    const destino = join(pasta, 'saida.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'export-pdf')
    await expect.poll(() => pdfText(destino), { timeout: 30_000 }).toMatch(/Página iii[\s\S]*Página iv/)
  })

  test('o campo inserido na faixa do Word vira campo no arquivo e conta na tela', async () => {
    await abrir()
    const peca = session.window.locator('.band--footer .band__text').first()
    await peca.click()
    await session.window.keyboard.press('Home')
    await session.window.getByRole('button', { name: 'Inserir total de páginas' }).click()
    await session.window.keyboard.type(' folhas — ')
    await session.window.locator('.ProseMirror').click()
    await expect.poll(() => rodapes(session)).toEqual(['2 folhas — Página iii', '2 folhas — Página iv'])

    const destino = join(pasta, 'saida.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const rodape = await entryOf(destino, 'word/footer1.xml')
    expect(rodape).toMatch(/w:instr=" NUMPAGES "/)
    expect(rodape).not.toContain('{total}')
  })

  test('primeira página diferente e o início mudados no painel valem na tela e no arquivo', async () => {
    await abrir()
    await menu(session, 'page-setup')
    const painel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    await painel.getByLabel('Primeira página diferente').check()
    await painel.getByLabel('Pares e ímpares diferentes').check()
    await painel.getByLabel('Formato do número').selectOption('upperRoman')
    await painel.getByLabel('Começar em').fill('1')
    await painel.getByRole('button', { name: 'Aplicar' }).click()

    // A title page without a footer; sheet 2 is page II, even, and without its own even footer.
    await expect.poll(() => rodapes(session)).toEqual([])

    const destino = join(pasta, 'saida.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const documento = await entryOf(destino, 'word/document.xml')
    expect(documento).toContain('<w:titlePg')
    expect(documento).toMatch(/<w:pgNumType w:fmt="upperRoman" w:start="1"/)
    expect(await entryOf(destino, 'word/settings.xml')).toContain('w:evenAndOddHeaders')
  })

  test('documento novo: o número de página inserido no rodapé aparece em cada folha', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda.')
    await session.window.getByRole('button', { name: 'Inserir número da página' }).click()
    await expect.poll(() => rodapes(session)).toEqual(['1', '2'])
  })
})
