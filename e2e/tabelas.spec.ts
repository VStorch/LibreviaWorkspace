/// <reference lib="dom" />
// O corpo de `evaluate` roda no renderer, mas é compilado no escopo do Node.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithStretchedImage, docxWithTable, entryOf } from './fixtures.js'

/**
 * Tabelas e imagens editáveis — o marco M4.
 *
 * Os comandos do TableKit existiam desde que a tabela passou a ser lida do
 * `.docx`, e não tinham interface nenhuma: a barra inseria uma 3 × 3 fixa e mais
 * nada. Aqui se confere o caminho que a pessoa usa — o menu "Tabela", o botão
 * direito dentro da célula, o diálogo de propriedades e as alças da imagem.
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
   * A tabela vem do arquivo, e não do comando de inserir, de propósito: o
   * histórico junta numa entrada só as mudanças vizinhas feitas em menos de meio
   * segundo, e uma tabela inserida logo antes da linha iria embora no mesmo
   * desfazer. A tabela lida não é transação nenhuma, então o desfazer só tem a
   * linha para voltar — que é o que este teste quer conferir.
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

    // O menu de contexto devolve o foco ao editor depois de fechar; o atalho só
    // chega ao histórico quando ele já voltou.
    await expect(session.window.locator('.page__content[contenteditable="true"]')).toBeFocused()

    // Uma ação, um desfazer: o comando do TableKit é uma transação só.
    await session.window.keyboard.press('Control+z')
    await expect.poll(linhas).toBe(2)
  })

  /**
   * O que se perde **ao gravar** aparece na mesma faixa do que se perde ao abrir.
   *
   * O sidecar registrava a perda, o processo main a devolvia, e o renderer só
   * olhava o inventário na abertura: a mesclagem vertical feita na tela sumia do
   * arquivo e a pessoa via "Salvo".
   */
  test('a perda na hora de salvar aparece na faixa de aviso', async () => {
    // Gravar em `.docx` pede um `.docx` de origem: é sobre ele que o sidecar
    // escreve. A tabela mesclada é inserida nele.
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

    // Muitos passos de propósito: se cada um virasse transação, o desfazer abaixo
    // voltaria só o último pixel.
    await session.window.mouse.move(caixa.x + 5, caixa.y + 5)
    await session.window.mouse.down()
    await session.window.mouse.move(caixa.x - 95, caixa.y + 5, { steps: 20 })
    await session.window.mouse.up()

    // Canto com a proporção travada: 400 × 100 vira 300 × 75.
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(300)
    const altura = Math.round((await imagem.boundingBox())?.height ?? 0)
    expect(altura).toBe(75)

    await session.window.keyboard.press('Control+z')
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(400)
  })

  test('arrastar a alça para fora não passa da largura da coluna', async () => {
    const origem = join(pasta, 'imagem.docx')
    await writeFile(origem, await docxWithStretchedImage())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const imagem = session.window.locator('.page__content .image-frame img')
    await expect(imagem).toBeVisible()
    await imagem.click()

    // Pelo teclado, e não arrastando: o ponteiro para na borda da janela, e o
    // tamanho dela decidiria se o teste chega ao teto. Cinquenta passos de oito
    // pixels passam de qualquer coluna.
    const alça = session.window.getByRole('button', { name: 'Redimensionar imagem pela borda da direita' })
    await alça.focus()
    for (let passo = 0; passo < 50; passo += 1) await session.window.keyboard.press('ArrowRight')

    // O teto é a coluna, medida no parágrafo — e não no embrulho em linha do
    // NodeView, cuja largura é zero e deixava o teto infinito. O CSS escondia o
    // excesso na tela, mas a medida pedida ia ao arquivo.
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
   * A imagem redimensionada continua no parágrafo dela, com o desenho dela.
   *
   * O editor tratava a imagem como bloco, e a que vem do `.docx` mora dentro de
   * um parágrafo. A primeira mudança de atributo não cabia ali: o ProseMirror
   * partia o parágrafo, a imagem descia para um novo e o original ficava vazio —
   * dois blocos reescritos, o `wp:docPr` "Quadrado" trocado por "Imagem 2", o
   * relacionamento trocado, e o desfazer sem efeito.
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

    // Pelo teclado: cada seta é um passo, e as alças continuam lá depois dele.
    const alça = session.window.getByRole('button', {
      name: 'Redimensionar imagem pelo canto inferior direito',
    })
    await alça.focus()
    await session.window.keyboard.press('ArrowLeft')
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(392)
    await expect(alça).toBeVisible()
    await session.window.keyboard.press('ArrowLeft')
    await expect.poll(async () => Math.round((await imagem.boundingBox())?.width ?? 0)).toBe(384)

    // A imagem não desce: nenhum parágrafo vazio apareceu em cima dela.
    expect(Math.round((await imagem.boundingBox())?.y ?? 0)).toBe(Math.round(topo))
    await expect(session.window.locator('.page__content p')).toHaveCount(2)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const antes = await entryOf(origem, 'word/document.xml')
    const xml = await entryOf(destino, 'word/document.xml')
    expect(xml.match(/<w:p[ >]/g)).toHaveLength(antes.match(/<w:p[ >]/g)?.length ?? -1)
    expect(xml).toMatch(/<wp:docPr id="1" name="Quadrado"\s*\/>/)
    expect(xml).toContain('r:embed="rId9"')
    // 384 × 96 px, na proporção 4 : 1 do arquivo.
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
    // O `w:jc` mora no parágrafo que tem a imagem, e o desenho é o de antes.
    expect(xml).toMatch(/<w:p><w:pPr>(?:(?!<\/w:pPr>).)*<w:jc w:val="center"\s*\/><\/w:pPr><w:r><w:drawing>/)
    expect(xml).toMatch(/<wp:docPr id="1" name="Quadrado" descr="Novo texto"\s*\/>/)

    // E um Ctrl+Z desfaz o gesto inteiro — texto e alinhamento juntos.
    await session.window.keyboard.press('Control+z')
    await expect(imagem).toHaveAttribute('alt', 'Original')
    await expect(session.window.locator('.page__content p').first()).not.toHaveCSS('text-align', 'center')
  })
})
