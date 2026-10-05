/// <reference lib="dom" />
// The `evaluate` body runs in the renderer, but is compiled in Node's scope.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithLongTable, docxWithStretchedImage, docxWithTable, entryOf } from './fixtures.js'

/**
 * Editable tables and images, along the user's path: the "Table" menu, right click on a cell, the
 * properties dialog and the image handles.
 */
test.describe('tabelas e imagens editáveis', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-tabelas-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('inserir tabela pergunta linhas e colunas', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.page__content').click()

    await menu(session, 'table-insert')
    const dialog = session.window.getByRole('dialog', { name: 'Inserir tabela' })
    await dialog.getByRole('spinbutton', { name: 'Linhas' }).fill('2')
    await dialog.getByRole('spinbutton', { name: 'Colunas' }).fill('4')
    await dialog.getByRole('button', { name: 'Inserir' }).click()

    const table = session.window.locator('.page__content table')
    await expect(table.locator('tr')).toHaveCount(2)
    await expect(table.locator('tr').first().locator('th, td')).toHaveCount(4)
  })

  /**
   * The table comes from the file: history merges neighbouring changes less than half a second
   * apart, and an inserted table would go away in the same undo as the row.
   */
  test('o botão direito dentro da tabela oferece as ações dela', async () => {
    const origem = join(pasta, 'tabela.docx')
    await writeFile(origem, await docxWithTable())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const linhas = (): Promise<number> =>
      session.window
        .locator('.page__content table')
        .first()
        .evaluate((table) => (table as HTMLTableElement).rows.length)

    const célula = session.window.locator('.page__content td', { hasText: 'Dado B' })
    await expect(célula).toBeVisible()
    await expect.poll(linhas).toBe(2)

    await célula.click({ button: 'right' })
    await session.window.getByRole('menuitem', { name: 'Inserir linha abaixo' }).click()
    await expect.poll(linhas).toBe(3)

    // The context menu returns focus to the editor after closing; the shortcut only reaches the
    // history once focus is back.
    await expect(session.window.locator('.page__content[contenteditable="true"]')).toBeFocused()

    // One action, one undo: the TableKit command is a single transaction.
    await session.window.keyboard.press('Control+z')
    await expect.poll(linhas).toBe(2)
  })

  /** What is lost on save shows in the same banner as what is lost on open. */
  test('a perda na hora de salvar aparece na faixa de aviso', async () => {
    // Saving to `.docx` needs a source `.docx`, where the merged table goes in.
    const origem = join(pasta, 'mesclada.docx')
    await writeFile(origem, await docxWithTable())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await session.window.getByText('Antes da tabela.').click()
    await session.window.keyboard.press('End')
    await menu(session, 'table-insert')
    await session.window
      .getByRole('dialog', { name: 'Inserir tabela' })
      .getByRole('button', { name: 'Inserir' })
      .click()

    const linhas = session.window.locator('.page__content table').first().locator('tr')
    await linhas.nth(1).locator('td').first().click()
    await linhas
      .nth(2)
      .locator('td')
      .first()
      .click({ modifiers: ['Shift'] })
    await linhas.nth(2).locator('td').first().click({ button: 'right' })
    await session.window.getByRole('menuitem', { name: 'Mesclar células' }).click()
    await expect(session.window.locator('.page__content td[rowspan="2"]')).toHaveCount(1)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const faixa = session.window.getByRole('status').filter({ hasText: 'mesclagem vertical' })
    await expect(faixa).toBeVisible()
    await expect(faixa).toContainText('Nesta gravação, isto não chegou ao arquivo')
  })

  test('o sombreamento escolhido nas propriedades aparece na célula', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.page__content').click()
    await menu(session, 'table-insert')
    await session.window
      .getByRole('dialog', { name: 'Inserir tabela' })
      .getByRole('button', { name: 'Inserir' })
      .click()

    await session.window.locator('.page__content td').first().click()
    await menu(session, 'table-properties')

    const dialog = session.window.getByRole('dialog', { name: 'Propriedades da tabela' })
    await dialog.getByRole('checkbox', { name: 'Sombreamento' }).check()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()

    const fundo = await session.window
      .locator('.page__content td')
      .first()
      .evaluate((cell) => getComputedStyle(cell).backgroundColor)
    expect(fundo).toBe('rgb(217, 217, 217)')
  })

  test('arrastar a alça redimensiona a imagem, e desfazer volta num passo só', async () => {
    const origem = join(pasta, 'imagem.docx')
    await writeFile(origem, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const imagem = session.window.locator('.page__content .image-frame img')
    await expect(imagem).toBeVisible()
    await imagem.click()

    const alça = session.window.getByRole('button', {
      name: 'Redimensionar imagem pelo canto inferior direito',
    })
    const caixa = await alça.boundingBox()
    if (caixa === null) throw new Error('a alça não apareceu')

    // Many steps: if each became a transaction, undo would only bring back the last.
    await session.window.mouse.move(caixa.x + 5, caixa.y + 5)
    await session.window.mouse.down()
    await session.window.mouse.move(caixa.x - 95, caixa.y + 5, { steps: 20 })
    await session.window.mouse.up()

    // A corner with the ratio locked: 400 × 100 becomes 300 × 75.
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(300)
    const altura = Math.round((await imagem.boundingBox())?.height ?? 0)
    expect(altura).toBe(75)

    await session.window.keyboard.press('Control+z')
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(400)
  })

  test('com zoom de 150 %, arrastar a alça grava o tamanho que a pessoa vê', async () => {
    const origem = join(pasta, 'imagem.docx')
    await writeFile(origem, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    for (let i = 0; i < 3; i++) await menu(session, 'zoom-in')
    await expect(session.window.locator('.statusbar__zoom-level')).toHaveText('150%')

    const imagem = session.window.locator('.page__content .image-frame img')
    await expect(imagem).toBeVisible()
    await imagem.click()
    const alça = session.window.getByRole('button', {
      name: 'Redimensionar imagem pelo canto inferior direito',
    })
    const caixa = await alça.boundingBox()
    if (caixa === null) throw new Error('a alça não apareceu')

    // 150 screen px at 150 % are 100 document px: 400 × 100 becomes 300 × 75, exactly like the same
    // gesture at 100 % with 100 px.
    await session.window.mouse.move(caixa.x + 5, caixa.y + 5)
    await session.window.mouse.down()
    await session.window.mouse.move(caixa.x - 145, caixa.y + 5, { steps: 20 })
    await session.window.mouse.up()

    const medida = () =>
      imagem.evaluate((img) => [
        (img as HTMLImageElement).offsetWidth,
        (img as HTMLImageElement).offsetHeight,
      ])
    await expect.poll(medida).toEqual([300, 75])
    // On screen, the saved image shows enlarged by the zoom.
    expect(Math.round((await imagem.boundingBox())!.width)).toBe(450)
  })

  test('com zoom de 150 %, arrastar a divisória da coluna muda a largura na escala do documento', async () => {
    const origem = join(pasta, 'tabela-longa.docx')
    await writeFile(origem, await docxWithLongTable(5))
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    for (let i = 0; i < 3; i++) await menu(session, 'zoom-in')
    await expect(session.window.locator('.statusbar__zoom-level')).toHaveText('150%')

    const célula = session.window.locator('.page__content td').first()
    await expect(célula).toBeVisible()
    const largura = () => célula.evaluate((cell) => (cell as HTMLElement).offsetWidth)
    const antes = await largura()
    const caixa = (await célula.boundingBox())!

    // The divider is the cell's right border; the plugin finds it on hover.
    const x = caixa.x + caixa.width - 2
    const y = caixa.y + caixa.height / 2
    await session.window.mouse.move(x - 20, y)
    await session.window.mouse.move(x, y, { steps: 5 })
    await session.window.mouse.down()
    await session.window.mouse.move(x - 150, y, { steps: 10 })
    await session.window.mouse.up()

    // 150 screen px at 150 % are 100 document px (150 without the conversion).
    await expect.poll(async () => antes - (await largura())).toBeGreaterThanOrEqual(98)
    expect(antes - (await largura())).toBeLessThanOrEqual(102)
  })

  test('arrastar a alça para fora não passa da largura da coluna', async () => {
    const origem = join(pasta, 'imagem.docx')
    await writeFile(origem, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const imagem = session.window.locator('.page__content .image-frame img')
    await expect(imagem).toBeVisible()
    await imagem.click()

    // By keyboard: the pointer stops at the window edge. Fifty steps of eight pixels pass any
    // column.
    const alça = session.window.getByRole('button', { name: 'Redimensionar imagem pela borda da direita' })
    await alça.focus()
    for (let passo = 0; passo < 50; passo += 1) await session.window.keyboard.press('ArrowRight')

    // The ceiling is the column measured on the paragraph, not on the NodeView's zero-width inline
    // wrapper.
    const medidas = await session.window.evaluate(() => {
      const img = document.querySelector('.page__content .image-frame img') as HTMLImageElement
      return {
        pedida: Number.parseFloat(img.style.width),
        coluna: (img.closest('p') as HTMLElement).clientWidth,
      }
    })
    expect(medidas.pedida).toBeGreaterThan(400)
    expect(medidas.pedida).toBeLessThanOrEqual(medidas.coluna)
  })

  /**
   * A `.docx` image lives in a paragraph: resizing it does not split it into another, nor replace
   * the `wp:docPr` and the relationship, and undo works.
   */
  test('redimensionar a imagem do arquivo não a tira do parágrafo', async () => {
    const origem = join(pasta, 'imagem.docx')
    const destino = join(pasta, 'saida.docx')
    await writeFile(origem, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')

    const imagem = session.window.locator('.page__content .image-frame img')
    await expect(imagem).toBeVisible()
    const topo = (await imagem.boundingBox())?.y ?? 0
    await imagem.click()

    // By keyboard: each arrow is a step, and the handles are still there after it.
    const alça = session.window.getByRole('button', {
      name: 'Redimensionar imagem pelo canto inferior direito',
    })
    await alça.focus()
    await session.window.keyboard.press('ArrowLeft')
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(392)
    await expect(alça).toBeVisible()
    await session.window.keyboard.press('ArrowLeft')
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(384)

    // The image does not move down: no empty paragraph appeared above it.
    expect(Math.round((await imagem.boundingBox())?.y ?? 0)).toBe(Math.round(topo))
    await expect(session.window.locator('.page__content p')).toHaveCount(2)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const antes = await entryOf(origem, 'word/document.xml')
    const xml = await entryOf(destino, 'word/document.xml')
    expect(xml.match(/<w:p[ >]/g)).toHaveLength(antes.match(/<w:p[ >]/g)?.length ?? -1)
    expect(xml).toMatch(/<wp:docPr id="1" name="Quadrado"\s*\/>/)
    expect(xml).toContain('r:embed="rId9"')
    // 384 × 96 px, at the file's 4:1 ratio.
    expect(xml).toMatch(/<wp:extent cx="3657600" cy="914400"\s*\/>/)
  })

  test('o alinhamento do diálogo vai para o parágrafo da imagem', async () => {
    const origem = join(pasta, 'imagem.docx')
    const destino = join(pasta, 'saida.docx')
    await writeFile(origem, await docxWithStretchedImage('Original'))
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')

    const imagem = session.window.locator('.page__content .image-frame img')
    await expect(imagem).toBeVisible()
    await imagem.click()
    await menu(session, 'image-properties')

    const dialog = session.window.getByRole('dialog', { name: 'Propriedades da imagem' })
    await dialog.getByRole('textbox', { name: 'Texto alternativo' }).fill('Novo texto')
    await dialog.getByRole('combobox', { name: 'Alinhamento da imagem' }).selectOption('center')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const antes = await entryOf(origem, 'word/document.xml')
    const xml = await entryOf(destino, 'word/document.xml')
    expect(xml.match(/<w:p[ >]/g)).toHaveLength(antes.match(/<w:p[ >]/g)?.length ?? -1)
    // `w:jc` lives on the paragraph holding the image, and the drawing is the previous one.
    expect(xml).toMatch(/<w:p><w:pPr>(?:(?!<\/w:pPr>).)*<w:jc w:val="center"\s*\/><\/w:pPr><w:r><w:drawing>/)
    expect(xml).toMatch(/<wp:docPr id="1" name="Quadrado" descr="Novo texto"\s*\/>/)

    // And one Ctrl+Z undoes the whole gesture, text and alignment together.
    await session.window.keyboard.press('Control+z')
    await expect(imagem).toHaveAttribute('alt', 'Original')
    await expect(session.window.locator('.page__content p').first()).not.toHaveCSS('text-align', 'center')
  })
})
