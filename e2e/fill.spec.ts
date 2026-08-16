import { expect, test, type Page } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * A alça de preenchimento leva a fórmula, não o resultado.
 *
 * Este é um teste de arraste de verdade, com o mouse: a lógica está coberta por
 * testes de unidade, mas eles não sabem se o evento certo do grid foi assinado.
 * `fillRange` existia e passava havia duas fases — só não tinha quem a chamasse.
 */
test.describe('alça de preenchimento', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('arrastar uma fórmula desloca as referências', async () => {
    await menu(session, 'new-spreadsheet')
    await expect(cell(session.window, 0, 0)).toBeVisible()

    await write(session.window, 0, 0, '2')
    await write(session.window, 1, 0, '3')
    await write(session.window, 0, 1, '10')
    await write(session.window, 1, 1, '20')
    await write(session.window, 0, 2, '=A1*B1')

    await expect(cell(session.window, 0, 2)).toHaveText('20')

    await dragHandle(session.window, 0, 2, 1, 2)

    // Copiar o texto exibido daria 20 aqui. A fórmula deslocada dá 60.
    await expect(cell(session.window, 1, 2)).toHaveText('60')

    await select(session.window, 1, 2)
    await expect(session.window.locator('.formula-bar__input')).toHaveValue('=A2*B2')
  })
})

function cell(window: Page, row: number, column: number) {
  return window
    .locator(`revogr-overlay-selection revogr-data [data-rgrow="${row}"][data-rgcol="${column}"]`)
    .first()
}

async function select(window: Page, row: number, column: number): Promise<void> {
  await cell(window, row, column).click()
  await expect(window.locator('.formula-bar__ref')).toHaveText(
    `${String.fromCharCode(65 + column)}${row + 1}`,
  )
}

async function write(window: Page, row: number, column: number, text: string): Promise<void> {
  await select(window, row, column)
  const input = window.locator('.formula-bar__input')
  await input.fill(text)
  await input.press('Enter')
}

/**
 * Arrasta o quadradinho do canto inferior direito da seleção.
 *
 * Ele fica na quina da célula selecionada, então o arraste começa alguns pixels
 * para dentro do canto — o mesmo gesto que o usuário faz.
 */
async function dragHandle(
  window: Page,
  fromRow: number,
  fromColumn: number,
  toRow: number,
  toColumn: number,
): Promise<void> {
  await select(window, fromRow, fromColumn)

  const origin = await cell(window, fromRow, fromColumn).boundingBox()
  const destination = await cell(window, toRow, toColumn).boundingBox()
  if (origin === null || destination === null) throw new Error('célula sem posição na tela')

  await window.mouse.move(origin.x + origin.width - 2, origin.y + origin.height - 2)
  await window.mouse.down()
  await window.mouse.move(destination.x + destination.width / 2, destination.y + destination.height / 2, {
    steps: 8,
  })
  await window.mouse.up()
}
