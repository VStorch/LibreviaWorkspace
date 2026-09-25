import { execFile } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithLongTable } from './fixtures.js'

/**
 * O editor pagina ao vivo.
 *
 * Até aqui a tela era uma tira contínua com marcas tracejadas de estimativa, e
 * quem paginava de verdade era só a exportação — a decisão registrada no §6.3 do
 * plano. Agora o texto corre sobre folhas desenhadas, e o número de folhas
 * responde ao que se digita.
 *
 * A contagem de folhas responde à digitação. Na tabela longa, também se
 * confere que a última linha está dentro do papel: contar folhas sozinho não
 * detecta conteúdo desenhado além da borda. Não se fixa a fonte nem a linha
 * exata do corte, que variam conforme a máquina.
 */
test.describe('paginação ao vivo', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-paginacao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('tabela longa termina dentro da última folha', async () => {
    const source = join(pasta, 'tabela-longa.docx')
    await writeFile(source, await docxWithLongTable())
    await stubDialogs(session.app, { open: source, messageBox: 1 })
    await menu(session, 'open')
    const sheets = session.window.locator('.paper')
    const rows = session.window.locator('.ProseMirror table').first().locator('tr')
    await expect(rows).toHaveCount(80)
    await expect.poll(() => sheets.count()).toBeGreaterThanOrEqual(2)
    await expect(async () => {
      const row = await rows.last().boundingBox()
      const sheet = await sheets.last().boundingBox()
      expect(row).not.toBeNull()
      expect(sheet).not.toBeNull()
      expect(row!.x).toBeGreaterThanOrEqual(sheet!.x)
      expect(row!.y).toBeGreaterThanOrEqual(sheet!.y)
      expect(row!.x + row!.width).toBeLessThanOrEqual(sheet!.x + sheet!.width)
      expect(row!.y + row!.height).toBeLessThanOrEqual(sheet!.y + sheet!.height)
    }).toPass()
  })

  test('documento novo tem uma folha só', async () => {
    await menu(session, 'new-document')
    await expect(session.window.locator('.paper')).toHaveCount(1)
    await expect(session.window.locator('.statusbar__metric', { hasText: 'página' })).toContainText(
      '1 página',
    )
  })

  test('a quebra de página pedida à mão abre folha nova', async () => {
    // O medidor anterior ignorava o nó `pageBreak`: um documento com capa e
    // sumário aparecia como uma folha só, por mais quebras que tivesse.
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Capa.')

    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda folha.')

    await expect(session.window.locator('.paper')).toHaveCount(2)
  })

  test('texto que não cabe empurra para a folha seguinte', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()

    // Parágrafos de sobra para estourar uma A4 com margens de 25 mm. Digitados
    // pelo teclado, e não injetados no modelo: é a digitação que precisa fazer
    // a folha nascer.
    for (let i = 0; i < 60; i++) {
      await session.window.keyboard.type(`Linha ${i} de um parágrafo qualquer para ocupar a folha.`)
      await session.window.keyboard.press('Enter')
    }

    await expect(session.window.locator('.paper')).not.toHaveCount(1)
  })

  test('cada folha repete o cabeçalho com o número dela', async () => {
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda.')

    // A numeração ao lado da folha é o que diz "isto é página 2", e antes daqui
    // toda faixa recebia `pageNumber={1}` fixo.
    await expect(session.window.locator('.paper__number')).toHaveText(['1', '2'])
  })

  test('o papel sai com as mesmas folhas que a tela mostra', async () => {
    // Havia dois paginadores que precisavam concordar: o nosso, na tela, e o do
    // Chromium, na exportação. Concordar por coincidência é o que este teste
    // recusa — agora o papel é montado a partir das folhas da tela, e o número
    // não pode divergir.
    const destino = join(pasta, 'saida.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })

    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Primeira folha.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Segunda folha.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.type('Terceira folha.')

    // Esperar as folhas, e não contá-las de uma vez: a paginação assenta num
    // quadro depois da última tecla, e `count()` não repete a pergunta. Era
    // esta linha que reprovava sob carga — a terceira folha ainda não existia —
    // e passava sozinha, onde a máquina tem folga de sobra.
    const folhas = session.window.locator('.paper')
    await expect(folhas).toHaveCount(3)
    const naTela = await folhas.count()

    await menu(session, 'export-pdf')
    await expect.poll(async () => contarPaginas(destino), { timeout: 30000 }).toBe(naTela)
  })

  test('o parágrafo que não cabe é cortado entre linhas, no mesmo lugar na tela e no PDF', async () => {
    test.skip(!(await temPdftotext()), 'pdftotext não instalado')
    const destino = join(pasta, 'corte.pdf')
    await stubDialogs(session.app, { save: destino, messageBox: 1 })
    await paragrafoAtravessandoAFolha(session)

    const folhas = session.window.locator('.paper')
    await expect(folhas).toHaveCount(2)
    await expect.poll(() => corteNaTela(session)).not.toBeNull()
    const corte = (await corteNaTela(session))!

    // Na tela: a linha antes do espaçador termina dentro da primeira folha, e a
    // de depois começa dentro da segunda.
    const limites = await session.window.evaluate((palavra) => {
      const paragraph = document.querySelector('.ProseMirror .page-line-gap')!.closest('p')!
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
        const at = ` ${node.textContent ?? ''} `.indexOf(` ${palavra} `)
        if (at === -1) continue
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + palavra.length)
        return range.getBoundingClientRect().top
      }
      return null
    }, corte.depois)
    const sheets = await folhas.evaluateAll((all) => all.map((sheet) => sheet.getBoundingClientRect().top))
    expect(limites).not.toBeNull()
    expect(limites!).toBeGreaterThan(sheets[1]!)

    await menu(session, 'export-pdf')
    await expect.poll(async () => contarPaginas(destino), { timeout: 30000 }).toBe(2)
    const primeira = await palavrasDaPagina(destino, 1)
    const segunda = await palavrasDaPagina(destino, 2)
    expect(primeira.at(-1)).toBe(corte.antes)
    expect(segunda[0]).toBe(corte.depois)
  })

  test('viúvas e órfãs: a órfã no pé leva o parágrafo inteiro para a folha seguinte', async () => {
    // Trinta parágrafos de enchimento deixam lugar para uma linha só no pé: sem
    // o controle, ela ficaria órfã; com ele, o parágrafo desce inteiro.
    await paragrafoAtravessandoAFolha(session, 30, 100)
    await expect(session.window.locator('.paper')).toHaveCount(2)
    await expect.poll(() => linhasEmVoltaDoCorte(session)).toEqual({ antes: 0, depois: 0 })
  })

  test('viúvas e órfãs: o corte deixa ao menos duas linhas de cada lado', async () => {
    // Vinte e oito: o parágrafo de cinco linhas cabe menos a última, e a quebra
    // leva junto a penúltima para ela não abrir a folha sozinha. Sessão própria,
    // e não um segundo documento na mesma: trocar de documento com o primeiro
    // sujo abre o aviso de descartar, e o texto caía no documento de antes.
    await paragrafoAtravessandoAFolha(session, 28, 100)
    // Em `toPass`, e não lido uma vez: sob carga a medida ainda assenta, e um
    // corte de passagem não é o que a folha termina mostrando.
    await expect(async () => {
      const linhas = await linhasEmVoltaDoCorte(session)
      expect(linhas.antes).toBeGreaterThanOrEqual(2)
      expect(linhas.depois).toBeGreaterThanOrEqual(2)
    }).toPass()
  })

  test('a linha de cabeçalho da tabela se repete no alto de cada folha, na tela e no PDF', async () => {
    test.skip(!(await temPdftotext()), 'pdftotext não instalado')
    const source = join(pasta, 'cabecalho.docx')
    const destino = join(pasta, 'cabecalho.pdf')
    await writeFile(source, await docxWithLongTable(80, true))
    await stubDialogs(session.app, { open: source, save: destino, messageBox: 1 })
    await menu(session, 'open')

    const sheets = session.window.locator('.paper')
    await expect.poll(() => sheets.count()).toBeGreaterThanOrEqual(2)
    const repetidos = session.window.locator('.page-repeated-header')
    await expect.poll(() => repetidos.count()).toBe((await sheets.count()) - 1)

    // A cópia fica dentro da segunda folha, acima da primeira linha dela.
    const copia = await repetidos.first().boundingBox()
    const folha = await sheets.nth(1).boundingBox()
    expect(copia!.y).toBeGreaterThanOrEqual(folha!.y)
    expect(copia!.y + copia!.height).toBeLessThanOrEqual(folha!.y + folha!.height)
    await expect(repetidos.first()).toContainText('Cabeçalho repetido')
    // E alinhada com a tabela original, coluna por coluna.
    const original = await session.window.locator('.page__content th').first().boundingBox()
    const clone = await repetidos.first().locator('th').first().boundingBox()
    expect(Math.abs(clone!.x - original!.x)).toBeLessThanOrEqual(1)
    expect(Math.abs(clone!.width - original!.width)).toBeLessThanOrEqual(1)
    // Dentro do vão da linha que abre a folha, e não sobre o texto dela.
    const vizinha = await session.window.evaluate(() => {
      const header = document.querySelector('.page-repeated-header')!
      const row = header.closest('tr')!
      return {
        copia: header.getBoundingClientRect().bottom,
        linha: row.cells[0]!.getBoundingClientRect().top,
      }
    })
    expect(vizinha.copia).toBeGreaterThan(vizinha.linha)

    await menu(session, 'export-pdf')
    await expect.poll(async () => contarPaginas(destino), { timeout: 30000 }).toBe(await sheets.count())
    const segunda = await palavrasDaPagina(destino, 2)
    expect(segunda.slice(0, 2)).toEqual(['Cabeçalho', 'repetido'])

    // Desligar a linha de cabeçalho pelo menu da tabela tira a repetição.
    await session.window
      .locator('.page__content th', { hasText: 'Cabeçalho repetido' })
      .first()
      .click({ button: 'right' })
    await session.window.getByRole('menuitem', { name: 'Linha de cabeçalho' }).click()
    await expect(repetidos).toHaveCount(0)
  })

  test('o zoom aumenta a folha sem mudar onde a página corta', async () => {
    await paragrafoAtravessandoAFolha(session)
    const folhas = session.window.locator('.paper')
    await expect(folhas).toHaveCount(2)
    await expect.poll(() => corteNaTela(session)).not.toBeNull()
    const corte = await corteNaTela(session)
    const largura = (await folhas.first().boundingBox())!.width

    const nivel = session.window.locator('.statusbar__zoom-level')
    await expect(nivel).toHaveText('100%')
    await session.window.getByRole('button', { name: 'Ampliar' }).click()
    await menu(session, 'zoom-in')
    await expect(nivel).toHaveText('125%')

    // A folha cresce na tela; a paginação, medida em 100 %, fica onde estava.
    await expect.poll(async () => (await folhas.first().boundingBox())!.width).toBeCloseTo(largura * 1.25, 0)
    await expect(folhas).toHaveCount(2)
    expect(await corteNaTela(session)).toEqual(corte)

    await menu(session, 'zoom-out')
    await expect(nivel).toHaveText('110%')
    await session.window.getByRole('button', { name: 'Ajustar à largura' }).click()
    await expect(session.window.getByRole('button', { name: 'Ajustar à largura' })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    expect(await corteNaTela(session)).toEqual(corte)

    await menu(session, 'zoom-reset')
    await expect(nivel).toHaveText('100%')
    await expect.poll(async () => (await folhas.first().boundingBox())!.width).toBeCloseTo(largura, 0)
  })
})

/**
 * Enche a folha com parágrafos curtos e depois escreve um parágrafo longo de
 * palavras numeradas (`p1 p2 …`), que atravessa o pé da folha. As palavras
 * únicas são o que deixa comparar o corte da tela com o do papel sem depender
 * da fonte da máquina.
 */
async function paragrafoAtravessandoAFolha(session: Session, enchimento = 28, total = 160): Promise<void> {
  await menu(session, 'new-document')
  await session.window.locator('.ProseMirror').click()
  for (let i = 0; i < enchimento; i++) {
    await session.window.keyboard.insertText(`Enchimento ${i}.`)
    await session.window.keyboard.press('Enter')
  }
  const palavras = Array.from({ length: total }, (_, index) => `p${index + 1}`).join(' ')
  await session.window.keyboard.insertText(palavras)
}

/**
 * Quantas linhas do parágrafo cortado ficam antes e depois do espaçador —
 * zero e zero quando nenhum parágrafo foi cortado.
 */
async function linhasEmVoltaDoCorte(session: Session): Promise<{ antes: number; depois: number }> {
  return session.window.evaluate(() => {
    const gap = document.querySelector('.ProseMirror .page-line-gap')
    const paragraph = gap?.closest('p')
    if (gap === null || gap === undefined || paragraph === null || paragraph === undefined) {
      return { antes: 0, depois: 0 }
    }
    const lines = (range: Range) =>
      new Set(
        Array.from(range.getClientRects())
          .filter((rect) => rect.height > 0 && rect.width > 0)
          .map((rect) => Math.round(rect.top)),
      ).size
    const before = document.createRange()
    before.setStart(paragraph, 0)
    before.setEndBefore(gap)
    const after = document.createRange()
    after.setStartAfter(gap)
    after.setEnd(paragraph, paragraph.childNodes.length)
    return { antes: lines(before), depois: lines(after) }
  })
}

/** A última palavra antes do espaçador e a primeira depois, lidas do DOM. */
async function corteNaTela(session: Session): Promise<{ antes: string; depois: string } | null> {
  return session.window.evaluate(() => {
    const gap = document.querySelector('.ProseMirror .page-line-gap')
    const paragraph = gap?.closest('p')
    if (gap === null || gap === undefined || paragraph === null || paragraph === undefined) return null
    const before = document.createRange()
    before.setStart(paragraph, 0)
    before.setEndBefore(gap)
    const after = document.createRange()
    after.setStartAfter(gap)
    after.setEnd(paragraph, paragraph.childNodes.length)
    const words = (range: Range) => range.toString().trim().split(/\s+/)
    return { antes: words(before).at(-1) ?? '', depois: words(after)[0] ?? '' }
  })
}

/** As palavras de uma página do PDF, pelo `pdftotext` do sistema. */
async function palavrasDaPagina(caminho: string, pagina: number): Promise<string[]> {
  const { stdout } = await promisify(execFile)('pdftotext', [
    '-f',
    String(pagina),
    '-l',
    String(pagina),
    '-layout',
    caminho,
    '-',
  ])
  return stdout.split(/\s+/).filter((word) => word.length > 0)
}

async function temPdftotext(): Promise<boolean> {
  try {
    await promisify(execFile)('pdftotext', ['-v'])
    return true
  } catch {
    return false
  }
}

/**
 * Páginas de um PDF, contando os objetos `/Type /Page`.
 *
 * Sem biblioteca: o dado está no arquivo em texto claro, e trazer um leitor de
 * PDF inteiro para contar páginas seria desproporcional. `[^s]` no fim separa
 * `/Page` de `/Pages`, que é o nó da árvore e apareceria uma vez a mais.
 */
async function contarPaginas(caminho: string): Promise<number> {
  try {
    const bytes = await readFile(caminho)
    return (bytes.toString('latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length
  } catch {
    // Arquivo ainda não escrito: o `poll` tenta de novo.
    return 0
  }
}
