import { expect, test } from '@playwright/test'
import { clickMenuItem, launch, menu, type Session } from './app.js'
import { configurarAssistente, startFakeServer, type FakeServer } from './fake-ai-server.js'
import { cell, gridReady, select, selectDown, write } from './sheet-grid.js'

test.describe('assistente de IA na planilha', () => {
  let session: Session
  let fake: FakeServer

  test.beforeEach(async () => {
    fake = await startFakeServer()
    session = await launch()
    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
    const linhas = [
      ['Produto', 'Preço'],
      ['Caneta', '3'],
      ['Lápis', '5'],
    ]
    for (const [linha, valores] of linhas.entries()) {
      for (const [coluna, valor] of valores.entries()) await write(session.window, linha, coluna, valor)
    }
    await configurarAssistente(session, fake.baseUrl)
  })

  test.afterEach(async () => {
    await session.close()
    await new Promise((resolve) => fake.server.close(resolve))
  })

  function painel() {
    return session.window.getByRole('complementary', { name: 'Assistente de IA' })
  }

  async function pedir(tarefa: string, pedido?: string): Promise<void> {
    await menu(session, 'ai-assistant')
    await painel().getByLabel('O que fazer').selectOption({ label: tarefa })
    if (pedido !== undefined) await painel().getByLabel('Pedido').fill(pedido)
    await painel().getByRole('button', { name: 'Gerar' }).click()
  }

  async function desfazer(): Promise<void> {
    await cell(session.window, 9, 5).click()
    await session.window.keyboard.press('Control+z')
  }

  test('depois de escrever no pedido, clicar numa célula leva o foco ao grid e Editar > Desfazer desfaz a planilha', async () => {
    await menu(session, 'ai-assistant')
    const pedido = painel().getByLabel('Pedido')
    await pedido.click()
    await session.window.keyboard.type('abc')
    await cell(session.window, 9, 5).click()
    await expect(pedido).not.toBeFocused()

    await clickMenuItem(session, 'Editar', 'Desfazer')
    await expect(cell(session.window, 2, 1)).toHaveText('')
    await expect(pedido).toHaveValue('abc')
  })

  test('criar fórmula mostra a prévia com o valor, insere e um Ctrl+Z desfaz', async () => {
    fake.resposta = ['```\n=SOMA(', 'B2:B3)\n```']
    await select(session.window, 3, 1)
    await pedir('Criar fórmula', 'some os preços')

    await expect(painel().getByLabel('Fórmula', { exact: true })).toHaveText('=SOMA(B2:B3)')
    await expect(painel().getByLabel('Valor em B4')).toContainText('8')
    const corpo = fake.corpos.at(-1) ?? ''
    expect(corpo).toContain('Sheet: Planilha1')
    expect(corpo).toContain('B2: 3')
    expect(corpo).toContain('some os preços')

    await painel().getByRole('button', { name: 'Inserir na célula' }).click()
    await expect(cell(session.window, 3, 1)).toHaveText('8')
    await select(session.window, 3, 1)
    await expect(session.window.locator('.formula-bar__input')).toHaveValue('=SOMA(B2:B3)')

    await desfazer()
    await expect(cell(session.window, 3, 1)).toHaveText('')
    await expect(cell(session.window, 2, 1)).toHaveText('5')
  })

  test('fórmula com função desconhecida mostra o problema e não deixa inserir', async () => {
    fake.resposta = ['=XLOOKUP(A1;B1:B9;C1:C9)']
    await select(session.window, 3, 1)
    await pedir('Criar fórmula', 'procure o preço')

    await expect(painel().getByRole('alert')).toContainText('XLOOKUP')
    await expect(painel().getByRole('button', { name: 'Inserir na célula' })).toBeDisabled()
  })

  test('preencher três linhas mostra a prévia, aplica valores e um Ctrl+Z desfaz', async () => {
    fake.resposta = ['["Escrita", ', '"=1+1", 7]']
    await write(session.window, 3, 0, 'Borracha')
    await selectDown(session.window, 1, 2, 2)
    await pedir('Preencher coluna', 'categoria do produto da coluna A')

    const previa = painel().getByRole('table', { name: 'Prévia' })
    await expect(previa.locator('tbody tr')).toHaveCount(3)
    await expect(previa.locator('tbody tr').first()).toHaveText(/C2\s*Escrita/)
    const corpo = fake.corpos.at(-1) ?? ''
    expect(corpo).toContain('Row 2: A2: Caneta | B2: 3')
    expect(corpo).toContain('Columns: A (Produto) | B (Preço)')

    await painel().getByRole('button', { name: 'Aplicar' }).click()
    await expect(cell(session.window, 1, 2)).toHaveText('Escrita')
    await expect(cell(session.window, 2, 2)).toHaveText('=1+1')
    await expect(cell(session.window, 3, 2)).toHaveText('7')

    await desfazer()
    for (const linha of [1, 2, 3]) await expect(cell(session.window, linha, 2)).toHaveText('')
    await expect(cell(session.window, 3, 0)).toHaveText('Borracha')
  })

  test('resposta de preenchimento fora do formato não deixa aplicar', async () => {
    fake.resposta = ['["só um"]']
    await selectDown(session.window, 1, 2, 1)
    await pedir('Preencher coluna', 'categoria')
    await expect(painel().getByRole('alert')).toHaveText('A resposta não veio no formato esperado.')
    await expect(painel().getByRole('button', { name: 'Aplicar' })).toHaveCount(0)
  })

  test('explicar fórmula aparece formatado', async () => {
    fake.resposta = ['**Soma** os ', 'preços.\n\n- B2\n- B3']
    await write(session.window, 3, 1, '=SOMA(B2:B3)')
    await select(session.window, 3, 1)
    await pedir('Explicar fórmula')

    const resposta = painel().getByLabel('Resposta')
    await expect(resposta.locator('strong')).toHaveText('Soma')
    await expect(resposta.locator('li')).toHaveCount(2)
    expect(fake.corpos.at(-1) ?? '').toContain('Formula: =SOMA(B2:B3)')
  })

  test('o comando do atalho abre e fecha o painel', async () => {
    await menu(session, 'ai-assistant')
    await expect(painel().getByRole('button', { name: 'Gerar' })).toBeVisible()
    await menu(session, 'ai-assistant')
    await expect(painel()).toHaveCount(0)
  })
})
