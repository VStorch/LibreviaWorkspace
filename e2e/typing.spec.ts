import { expect, test, type Page } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * Lançar uma coluna de números sem parar entre eles.
 *
 * O grid espera 70 ms fixos depois do Enter antes de mover o foco para baixo, e
 * quem digita continuadamente acerta essa janela: a tecla chegava enquanto ele
 * ainda apontava para a célula anterior, e `1200` abaixo de `980` virava `200`.
 * Perda silenciosa — o erro só aparece quando a soma não bate.
 *
 * Os testes de unidade não alcançam isto: o defeito não está em nenhuma função
 * nossa, está no encontro do relógio do grid com o de quem digita. Por isso
 * aqui, com teclado de verdade e o intervalo de uma digitação normal.
 *
 * A entrada pela barra de fórmulas — o caminho que os outros testes usam — não
 * passava por essa janela, que é o motivo de o defeito ter durado tanto.
 */
test.describe('digitação contínua na planilha', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('uma coluna digitada sem pausa chega inteira', async () => {
    await menu(session, 'new-spreadsheet')
    await expect(cell(session.window, 0, 0)).toBeVisible()

    const valores = ['980', '1200', '2450', '860']
    await cell(session.window, 0, 0).click()

    for (const valor of valores) {
      // 60 ms por tecla é digitação rápida de teclado numérico, e nenhuma pausa
      // entre o Enter e o número seguinte — que é como se lança uma coluna.
      await session.window.keyboard.type(valor, { delay: 60 })
      await session.window.keyboard.press('Enter')
    }

    for (const [linha, valor] of valores.entries()) {
      await expect(cell(session.window, linha, 0)).toHaveText(valor)
    }
  })
})

function cell(window: Page, row: number, column: number) {
  return window
    .locator(`revogr-overlay-selection revogr-data [data-rgrow="${row}"][data-rgcol="${column}"]`)
    .first()
}
