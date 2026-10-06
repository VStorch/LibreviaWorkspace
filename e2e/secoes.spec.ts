import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithColumns, docxWithSections, entryOf } from './fixtures.js'
import { hasPdfinfo } from './external-tools.js'

/**
 * Sections: each sheet with its section's paper, band and number.
 *
 * The document has portrait (Roman), landscape (restarting at 1) and portrait again starting on an
 * odd page; Word inserts a blank sheet before it, already with the new section's paper. The last
 * two sections declare no footer, and inherit the first one's.
 */
test.describe('seções', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-secoes-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  async function abrir(): Promise<void> {
    const origem = join(pasta, 'secoes.docx')
    await writeFile(origem, await docxWithSections())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(4)
  }

  test('retrato e paisagem na mesma pilha, com a folha em branco da seção ímpar', async () => {
    await abrir()
    const folhas = await session.window
      .locator('.paper')
      .evaluateAll((papeis) => papeis.map((papel) => papel.getBoundingClientRect()))

    // The second sheet is wider than tall; the others, portrait.
    expect(folhas[0]!.width).toBeLessThan(folhas[0]!.height)
    expect(folhas[1]!.width).toBeGreaterThan(folhas[1]!.height)
    expect(folhas[2]!.width).toBeLessThan(folhas[2]!.height)
    // Centered in the stack: the portrait one starts further right than the landscape one.
    expect(folhas[0]!.left).toBeGreaterThan(folhas[1]!.left)

    // The landscape section's text takes its width, not the portrait one's.
    const paisagem = session.window.locator('.ProseMirror p', { hasText: 'Folha em paisagem.' })
    const caixa = await paisagem.boundingBox()
    expect(caixa!.x).toBeLessThan(folhas[0]!.left)
    expect(caixa!.y).toBeGreaterThan(folhas[1]!.top)
    expect(caixa!.y).toBeLessThan(folhas[1]!.bottom)

    // The third sheet is the blank one; portrait again opens the fourth.
    const retrato = await session.window
      .locator('.ProseMirror p', { hasText: 'Retrato outra vez' })
      .boundingBox()
    expect(retrato!.y).toBeGreaterThan(folhas[3]!.top)
  })

  test('o rodapé herdado numera cada seção com o formato e o reinício dela', async () => {
    await abrir()
    await expect
      .poll(() =>
        session.window
          .locator('.band--footer')
          .evaluateAll((faixas) => faixas.map((faixa) => (faixa.textContent ?? '').trim())),
      )
      .toEqual(['Página i', 'Página 1', 'Página 2', 'Página 3'])
  })

  test('o PDF sai com o papel de cada seção e os mesmos números', async () => {
    test.skip(!(await hasPdfinfo()), 'pdfinfo não instalado')
    await abrir()
    const destino = join(pasta, 'secoes.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'export-pdf')

    await expect.poll(() => tamanhos(destino), { timeout: 30_000 }).toHaveLength(4)
    const [primeira, segunda, terceira, quarta] = await tamanhos(destino)
    expect(primeira).toEqual('retrato')
    expect(segunda).toEqual('paisagem')
    expect(terceira).toEqual('retrato')
    expect(quarta).toEqual('retrato')

    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', destino, '-'])
    expect(stdout).toMatch(/Página i[\s\S]*Folha em paisagem[\s\S]*Página 1[\s\S]*Página 2[\s\S]*Página 3/)
  })

  const orientacoes = (session: Session) =>
    session.window.locator('.paper').evaluateAll((papeis) =>
      papeis.map((papel) => {
        const caixa = papel.getBoundingClientRect()
        return caixa.width > caixa.height ? 'paisagem' : 'retrato'
      }),
    )

  test('inserir quebra de seção, virar só a seção de baixo e excluir a quebra', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Em retrato.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Em paisagem.')
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-section-next-page')
    await expect(session.window.locator('.paper')).toHaveCount(2)

    // The cursor is in the lower section: "this section" means only that one.
    await session.window.locator('.ProseMirror p', { hasText: 'Em paisagem.' }).click()
    await menu(session, 'page-setup')
    const painel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    await painel.getByLabel('Orientação').selectOption('landscape')
    await expect(painel.getByLabel('Nesta seção')).toBeChecked()
    await painel.getByRole('button', { name: 'Aplicar' }).click()
    await expect.poll(() => orientacoes(session)).toEqual(['retrato', 'paisagem'])

    const destino = join(pasta, 'duas-secoes.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo.match(/<w:sectPr/g)).toHaveLength(2)
    expect(corpo).toMatch(/<w:p>(?:(?!<\/w:p>).)*<w:sectPr(?:(?!<\/w:sectPr>).)*w:h="16838"/)
    expect(corpo).toMatch(/<w:sectPr(?:(?!<\/w:sectPr>).)*w:orient="landscape"(?:(?!<w:sectPr).)*<\/w:body>/)

    // With the break deleted, the range above takes the format of the one below, as in Word.
    await menu(session, 'delete-section-break')
    await expect.poll(() => orientacoes(session)).toEqual(['paisagem'])
  })

  test('desvincular o rodapé da seção de paisagem dá a ela um rodapé próprio', async () => {
    await abrir()
    await session.window.locator('.ProseMirror p', { hasText: 'Folha em paisagem.' }).click()
    await menu(session, 'page-setup')
    const painel = session.window.getByRole('dialog', { name: 'Configuração de página' })
    const vinculo = painel.getByLabel('Rodapé: vincular ao anterior')
    await expect(vinculo).toBeChecked()
    await vinculo.uncheck()
    await painel.getByRole('button', { name: 'Aplicar' }).click()

    // The copy is the landscape section's: editing it does not change the first one's footer.
    const peca = session.window.locator('.band--footer .band__text').nth(1)
    await peca.click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Anexo — ')
    await session.window.locator('.ProseMirror').click()
    await expect
      .poll(() =>
        session.window
          .locator('.band--footer')
          .evaluateAll((faixas) => faixas.map((faixa) => (faixa.textContent ?? '').trim())),
      )
      .toEqual(['Página i', 'Anexo — Página 1', 'Anexo — Página 2', 'Anexo — Página 3'])

    const destino = join(pasta, 'desvinculado.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(destino, 'word/footer1.xml')).not.toContain('Anexo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo.match(/<w:footerReference/g)).toHaveLength(2)
  })

  test('colunas equilibradas antes da seção contínua, na tela e no PDF', async () => {
    const origem = join(pasta, 'colunas.docx')
    await writeFile(origem, await docxWithColumns())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.paper')).toHaveCount(1)

    const caixa = (texto: string) =>
      session.window.locator('.ProseMirror p', { hasText: texto }).boundingBox()
    // Six paragraphs in two balanced columns: three on each side.
    await expect
      .poll(async () => (await caixa('Parágrafo 4'))!.x)
      .toBeGreaterThan((await caixa('Parágrafo 1'))!.x + 100)
    const primeiro = (await caixa('Parágrafo 1'))!
    const quarto = (await caixa('Parágrafo 4'))!
    expect(Math.abs(quarto.y - primeiro.y)).toBeLessThan(2)
    expect(quarto.width).toBeLessThan(primeiro.width + 1)
    // The lower section's text moves down to the foot of the tallest column, and takes the whole
    // sheet.
    const terceiro = (await caixa('Parágrafo 3'))!
    const depois = (await caixa('Depois das colunas'))!
    expect(depois.y).toBeGreaterThan(terceiro.y + terceiro.height - 1)
    expect(depois.width).toBeGreaterThan(primeiro.width * 1.5)
    await expect(session.window.locator('.paper-column-line')).toHaveCount(1)

    test.skip(!(await hasPdfinfo()), 'pdftotext não instalado')
    const destino = join(pasta, 'colunas.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'export-pdf')
    await expect.poll(() => tamanhos(destino), { timeout: 30_000 }).toHaveLength(1)
    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', destino, '-'])
    expect(stdout).toMatch(/Parágrafo 1 em colunas\.\s+Parágrafo 4 em colunas\./)
  })

  test('Formatar → Colunas e a quebra de coluna num documento novo', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Esquerda.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Direita.')
    await menu(session, 'format-columns')
    const dialogo = session.window.getByRole('dialog', { name: 'Colunas' })
    await dialogo.getByLabel('Número de colunas').fill('2')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()

    await session.window.locator('.ProseMirror p', { hasText: 'Esquerda.' }).click()
    await session.window.keyboard.press('End')
    await menu(session, 'insert-column-break')
    const esquerda = session.window.locator('.ProseMirror p', { hasText: 'Esquerda.' })
    const direita = session.window.locator('.ProseMirror p', { hasText: 'Direita.' })
    await expect
      .poll(async () => (await direita.boundingBox())!.x - (await esquerda.boundingBox())!.x)
      .toBeGreaterThan(100)

    const destino = join(pasta, 'colunas-novo.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toMatch(/<w:cols [^>]*w:num="2"/)
    expect(corpo).toContain('w:type="column"')
  })

  test('desfazer a quebra de seção desfaz a seção, na tela e no arquivo', async () => {
    // The section lives in the store library, and the text says whether it counts: undo removes the
    // mark and, with it, the section and the start the lower one gained.
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Antes.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Depois.')
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-section-odd-page')
    await expect(session.window.locator('.paper')).toHaveCount(3)

    await session.window.keyboard.press('Control+z')
    await expect(session.window.locator('.paper')).toHaveCount(1)

    const destino = join(pasta, 'desfeito.docx')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo.match(/<w:sectPr/g)).toHaveLength(1)
    expect(corpo).not.toContain('oddPage')
  })

  test('o rascunho de antes das seções recusa quebra e colunas dizendo por quê', async () => {
    // Saved along the old path, it would not take the break to the file: the change would show on
    // screen and vanish from the .docx.
    const rascunho = join(pasta, 'antigo.sdoc')
    await writeFile(
      rascunho,
      JSON.stringify({
        format: 'sdoc',
        version: 5,
        page: { size: 'A4', orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 } },
        doc: {
          type: 'doc',
          content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Rascunho antigo.' }] }],
        },
      }),
    )
    await stubDialogs(session.app, { open: rascunho, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Rascunho antigo.')
    await session.window.locator('.ProseMirror').click()

    await menu(session, 'insert-section-next-page')
    await expect(session.window.getByRole('alert')).toContainText('versão anterior')
    await expect(session.window.locator('.paper')).toHaveCount(1)
    await expect(session.window.locator('.ProseMirror p')).toHaveCount(1)
  })

  test('a quebra de seção num item de lista fica no item, e a lista continua na folha seguinte', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('- Primeiro item.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Segundo item.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Terceiro item.')
    await session.window.locator('.ProseMirror li', { hasText: 'Segundo item.' }).click()
    await menu(session, 'insert-section-next-page')
    await expect(session.window.locator('.paper')).toHaveCount(2)

    const folhas = await session.window
      .locator('.paper')
      .evaluateAll((papeis) => papeis.map((papel) => papel.getBoundingClientRect().top))
    const segundo = await session.window
      .locator('.ProseMirror li', { hasText: 'Segundo item.' })
      .boundingBox()
    const terceiro = await session.window
      .locator('.ProseMirror li', { hasText: 'Terceiro item.' })
      .boundingBox()
    expect(segundo!.y).toBeLessThan(folhas[1]!)
    expect(terceiro!.y).toBeGreaterThan(folhas[1]!)
  })
})

/** Each PDF sheet's orientation, by the size `pdfinfo` reports. */
async function tamanhos(caminho: string): Promise<string[]> {
  try {
    const { stdout } = await promisify(execFile)('pdfinfo', ['-f', '1', '-l', '99', caminho])
    return [...stdout.matchAll(/Page\s+\d+ size:\s+([\d.]+) x ([\d.]+)/g)].map((medida) =>
      Number(medida[1]) > Number(medida[2]) ? 'paisagem' : 'retrato',
    )
  } catch {
    return []
  }
}
