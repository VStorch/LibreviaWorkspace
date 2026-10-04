import { expect, test, type Page } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * Lançar uma coluna de números sem parar: o grid espera 70 ms depois do Enter para
 * descer o foco, e a tecla que chega nessa janela iria à célula anterior (`1200`
 * abaixo de `980` vira `200`). Está no encontro do relógio do grid com o de quem
 * digita, que teste de unidade não alcança. A barra de fórmulas não passa por ela.
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
