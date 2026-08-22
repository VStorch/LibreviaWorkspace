import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * O editor pagina ao vivo.
 *
 * Até aqui a tela era uma tira contínua com marcas tracejadas de estimativa, e
 * quem paginava de verdade era só a exportação — a decisão registrada no §6.3 do
 * plano. Agora o texto corre sobre folhas desenhadas, e o número de folhas
 * responde ao que se digita.
 *
 * Os testes olham a **contagem de folhas**, e não pixels: onde exatamente a
 * linha cai depende da fonte que a máquina tem, e um teste preso a isso reprova
 * por motivo errado. Quantas folhas o documento tem é a pergunta que a pessoa
 * faz, e é estável.
 */
test.describe('paginação ao vivo', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
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
})
