import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
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
})

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
