import { expect, test } from '@playwright/test'
import { clickMenuItem, launch, menu, type Session } from './app.js'
import { configureAssistant, startFakeServer, type FakeServer } from './fake-ai-server.js'
import { cell, gridReady, select, selectDown, write } from './sheet-grid.js'

test.describe('AI assistant in the spreadsheet', () => {
  let session: Session
  let fake: FakeServer

  test.beforeEach(async () => {
    fake = await startFakeServer()
    session = await launch()
    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
    const lines = [
      ['Produto', 'Preço'],
      ['Caneta', '3'],
      ['Lápis', '5'],
    ]
    for (const [line, values] of lines.entries()) {
      for (const [column, value] of values.entries()) await write(session.window, line, column, value)
    }
    await configureAssistant(session, fake.baseUrl)
  })

  test.afterEach(async () => {
    await session.close()
    await new Promise((resolve) => fake.server.close(resolve))
  })

  function panel() {
    return session.window.getByRole('complementary', { name: 'Assistente de IA' })
  }

  async function ask(task: string, request?: string): Promise<void> {
    await menu(session, 'ai-assistant')
    await panel().getByLabel('O que fazer').selectOption({ label: task })
    if (request !== undefined) await panel().getByLabel('Pedido').fill(request)
    await panel().getByRole('button', { name: 'Gerar' }).click()
  }

  async function undo(): Promise<void> {
    await cell(session.window, 9, 5).click()
    await session.window.keyboard.press('Control+z')
  }

  test('after typing in the request, clicking a cell moves focus to the grid and Edit > Undo undoes the spreadsheet', async () => {
    await menu(session, 'ai-assistant')
    const request = panel().getByLabel('Pedido')
    await request.click()
    await session.window.keyboard.type('abc')
    await cell(session.window, 9, 5).click()
    await expect(request).not.toBeFocused()

    await clickMenuItem(session, 'Editar', 'Desfazer')
    await expect(cell(session.window, 2, 1)).toHaveText('')
    await expect(request).toHaveValue('abc')
  })

  test('create formula shows the preview with the value, inserts it, and one Ctrl+Z undoes it', async () => {
    fake.reply = ['```\n=SOMA(', 'B2:B3)\n```']
    await select(session.window, 3, 1)
    await ask('Criar fórmula', 'some os preços')

    await expect(panel().getByLabel('Fórmula', { exact: true })).toHaveText('=SOMA(B2:B3)')
    await expect(panel().getByLabel('Valor em B4')).toContainText('8')
    const body = fake.bodies.at(-1) ?? ''
    expect(body).toContain('Sheet: Planilha1')
    expect(body).toContain('B2: 3')
    expect(body).toContain('some os preços')

    await panel().getByRole('button', { name: 'Inserir na célula' }).click()
    await expect(cell(session.window, 3, 1)).toHaveText('8')
    await select(session.window, 3, 1)
    await expect(session.window.locator('.formula-bar__input')).toHaveValue('=SOMA(B2:B3)')

    await undo()
    await expect(cell(session.window, 3, 1)).toHaveText('')
    await expect(cell(session.window, 2, 1)).toHaveText('5')
  })

  test('a formula with an unknown function shows the problem and cannot be inserted', async () => {
    fake.reply = ['=XLOOKUP(A1;B1:B9;C1:C9)']
    await select(session.window, 3, 1)
    await ask('Criar fórmula', 'procure o preço')

    await expect(panel().getByRole('alert')).toContainText('XLOOKUP')
    await expect(panel().getByRole('button', { name: 'Inserir na célula' })).toBeDisabled()
  })

  test('filling three rows shows the preview, applies the values, and one Ctrl+Z undoes it', async () => {
    fake.reply = ['["Escrita", ', '"=1+1", 7]']
    await write(session.window, 3, 0, 'Borracha')
    await selectDown(session.window, 1, 2, 2)
    await ask('Preencher coluna', 'categoria do produto da coluna A')

    const preview = panel().getByRole('table', { name: 'Prévia' })
    await expect(preview.locator('tbody tr')).toHaveCount(3)
    await expect(preview.locator('tbody tr').first()).toHaveText(/C2\s*Escrita/)
    const body = fake.bodies.at(-1) ?? ''
    expect(body).toContain('Row 2: A2: Caneta | B2: 3')
    expect(body).toContain('Columns: A (Produto) | B (Preço)')

    await panel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(cell(session.window, 1, 2)).toHaveText('Escrita')
    await expect(cell(session.window, 2, 2)).toHaveText('=1+1')
    await expect(cell(session.window, 3, 2)).toHaveText('7')

    await undo()
    for (const line of [1, 2, 3]) await expect(cell(session.window, line, 2)).toHaveText('')
    await expect(cell(session.window, 3, 0)).toHaveText('Borracha')
  })

  test('a fill reply in the wrong format cannot be applied', async () => {
    fake.reply = ['["só um"]']
    await selectDown(session.window, 1, 2, 1)
    await ask('Preencher coluna', 'categoria')
    await expect(panel().getByRole('alert')).toHaveText('A resposta não veio no formato esperado.')
    await expect(panel().getByRole('button', { name: 'Aplicar' })).toHaveCount(0)
  })

  test('explain formula shows formatted', async () => {
    fake.reply = ['**Soma** os ', 'preços.\n\n- B2\n- B3']
    await write(session.window, 3, 1, '=SOMA(B2:B3)')
    await select(session.window, 3, 1)
    await ask('Explicar fórmula')

    const reply = panel().getByLabel('Resposta')
    await expect(reply.locator('strong')).toHaveText('Soma')
    await expect(reply.locator('li')).toHaveCount(2)
    expect(fake.bodies.at(-1) ?? '').toContain('Formula: =SOMA(B2:B3)')
  })

  test('the shortcut command opens and closes the panel', async () => {
    await menu(session, 'ai-assistant')
    await expect(panel().getByRole('button', { name: 'Gerar' })).toBeVisible()
    await menu(session, 'ai-assistant')
    await expect(panel()).toHaveCount(0)
  })
})
