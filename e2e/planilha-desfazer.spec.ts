import { expect, test } from '@playwright/test'
import { clickMenuItem, launch, menu, type Session } from './app.js'
import { cell, gridReady, reference, select, selectDown, write } from './sheet-grid.js'

test.describe('desfazer na planilha', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('Ctrl+Z desfaz, Ctrl+Y e Ctrl+Shift+Z refazem, e desfazer não grava passo novo', async () => {
    const { window } = session
    await write(window, 0, 0, 'a')
    await write(window, 0, 1, 'b')

    // Focus is still in the formula bar, with nothing typed: the key belongs to the sheet.
    await window.keyboard.press('Control+z')
    await expect(cell(window, 0, 1)).toHaveText('')
    await window.keyboard.press('Control+y')
    await expect(cell(window, 0, 1)).toHaveText('b')

    await cell(window, 5, 3).click()
    await window.keyboard.press('Control+z')
    await expect(cell(window, 0, 1)).toHaveText('')
    await window.keyboard.press('Control+z')
    await expect(cell(window, 0, 0)).toHaveText('')
    await window.keyboard.press('Control+Shift+z')
    await expect(cell(window, 0, 0)).toHaveText('a')
    await window.keyboard.press('Control+y')
    await expect(cell(window, 0, 1)).toHaveText('b')
    await window.keyboard.press('Control+y')
    await expect(cell(window, 0, 1)).toHaveText('b')
  })

  test('depois de desfazer, a célula ativa mostra o valor de antes', async () => {
    const { window } = session
    await write(window, 0, 1, 'b')
    await select(window, 0, 1)
    await window.keyboard.press('Control+z')
    await expect(cell(window, 0, 1)).toHaveText('')
    await expect(window.locator('.formula-bar__input')).toHaveValue('')
  })

  test('desfazer até o que estava salvo tira o aviso de alteração', async () => {
    const { window } = session
    const title = await window.title()
    await write(window, 0, 0, 'a')
    await expect(window).not.toHaveTitle(title)
    await window.keyboard.press('Control+z')
    await expect(window).toHaveTitle(title)
  })

  test('Editar > Desfazer e Refazer do menu usam o histórico da planilha', async () => {
    const { window } = session
    await write(window, 0, 0, 'a')
    await cell(window, 5, 3).click()
    await clickMenuItem(session, 'Editar', 'Desfazer')
    await expect(cell(window, 0, 0)).toHaveText('')
    await clickMenuItem(session, 'Editar', 'Refazer')
    await expect(cell(window, 0, 0)).toHaveText('a')
  })

  test('com um diálogo aberto, Ctrl+Z não desfaz a planilha por trás', async () => {
    const { window } = session
    await write(window, 0, 0, 'a')
    await cell(window, 5, 3).click()
    await menu(session, 'document-properties')
    await expect(window.getByRole('dialog')).toBeVisible()
    await window.keyboard.press('Control+z')
    await window.getByRole('dialog').getByRole('button', { name: 'Cancelar' }).click()
    await expect(window.getByRole('dialog')).toHaveCount(0)
    await expect(cell(window, 0, 0)).toHaveText('a')
  })

  // The menu's accelerator does not hear keys sent by Playwright: the test checks that the sheet
  // leaves the key alone, so it reaches the menu.
  test('Ctrl+Shift+Y não é refazer: a planilha deixa a tecla para o atalho do assistente', async () => {
    const { window } = session
    await write(window, 0, 0, 'a')
    await cell(window, 5, 3).click()
    await window.keyboard.press('Control+z')
    await expect(cell(window, 0, 0)).toHaveText('')
    await window.evaluate(() => {
      globalThis.addEventListener('keydown', (event) => {
        if (event.key.toLowerCase() === 'y') document.body.dataset['taken'] = String(event.defaultPrevented)
      })
    })
    await window.keyboard.press('Control+Shift+Y')
    await expect(window.locator('body')).toHaveAttribute('data-taken', 'false')
    await expect(cell(window, 0, 0)).toHaveText('')
  })

  test('texto que começa com "=" continua texto ao ser editado de novo', async () => {
    const { window } = session
    await write(window, 0, 0, "'=1+1")
    await expect(cell(window, 0, 0)).toHaveText('=1+1')
    await select(window, 0, 0)
    await expect(window.locator('.formula-bar__input')).toHaveValue("'=1+1")
    await window.locator('.formula-bar__input').press('End')
    await window.locator('.formula-bar__input').press('Enter')
    await expect(cell(window, 0, 0)).toHaveText('=1+1')
  })
})

test.describe('seleção da planilha', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
  })

  test.afterEach(async () => {
    await session.close()
  })

  async function selectBlock(): Promise<void> {
    const { window } = session
    for (const [row, column] of [
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
      [2, 0],
      [2, 1],
    ] as const) {
      await write(window, row, column, `${reference(row, column)}`)
    }
    await expect(async () => {
      await selectDown(window, 0, 0, 2)
      await window.keyboard.press('Shift+ArrowRight')
      await expect(window.locator('.formula-bar__ref')).toHaveText('A1:B3', { timeout: 1000 })
    }).toPass({ timeout: 15_000 })
  }

  test('clique direito dentro da seleção mantém a seleção e o menu age nela toda', async () => {
    const { window } = session
    await selectBlock()
    await cell(window, 1, 1).click({ button: 'right' })
    await expect(window.locator('.formula-bar__ref')).toHaveText('A1:B3')
    await expect(window.getByRole('menuitem', { name: 'Excluir 3 linhas' })).toBeVisible()
    await window.getByRole('menuitem', { name: 'Limpar conteúdo' }).click()
    for (const [row, column] of [
      [0, 0],
      [1, 1],
      [2, 1],
    ] as const) {
      await expect(cell(window, row, column)).toHaveText('')
    }
  })

  test('clique direito fora da seleção vai para a célula clicada', async () => {
    const { window } = session
    await selectBlock()
    await cell(window, 5, 3).click({ button: 'right' })
    await expect(window.locator('.formula-bar__ref')).toHaveText('D6')
    await expect(window.getByRole('menuitem', { name: 'Excluir linha', exact: true })).toBeVisible()
  })

  test('com colunas congeladas, a célula e o intervalo selecionados têm a referência certa', async () => {
    const { window } = session
    await select(window, 2, 2)
    await window.getByRole('button', { name: 'Congelar até a seleção' }).click()

    await expect(async () => {
      await cell(window, 4, 3).click()
      await expect(window.locator('.formula-bar__ref')).toHaveText('D5', { timeout: 1000 })
      await window.keyboard.press('Shift+ArrowRight')
      await window.keyboard.press('Shift+ArrowDown')
      await expect(window.locator('.formula-bar__ref')).toHaveText('D5:E6', { timeout: 1000 })
    }).toPass({ timeout: 15_000 })

    await cell(window, 1, 1).click()
    await expect(window.locator('.formula-bar__ref')).toHaveText('B2')
  })

  test('com colunas congeladas, Shift+seta e Shift+clique estendem a seleção através da divisa', async () => {
    const { window } = session
    const ref = window.locator('.formula-bar__ref')
    await select(window, 2, 2)
    await window.getByRole('button', { name: 'Congelar até a seleção' }).click()

    // Repeated as `selectDown`; from another cell, so the retry starts over from a single cell.
    await expect(async () => {
      await select(window, 3, 3)
      await select(window, 0, 0)
      for (let step = 0; step < 3; step++) await window.keyboard.press('Shift+ArrowRight')
      await expect(ref).toHaveText('A1:D1', { timeout: 1000 })
    }).toPass({ timeout: 15_000 })
    // Each pane draws its part of the range.
    await expect(window.locator('revogr-overlay-selection .selection-border-range')).toHaveCount(2)

    await window.keyboard.press('Shift+ArrowDown')
    await expect(ref).toHaveText('A1:D2')
    await window.keyboard.press('Shift+ArrowLeft')
    await window.keyboard.press('Shift+ArrowLeft')
    await expect(ref).toHaveText('A1:B2')
    await expect(window.locator('revogr-overlay-selection .selection-border-range')).toHaveCount(1)

    await cell(window, 0, 0).click()
    await expect(ref).toHaveText('A1')
    await cell(window, 5, 4).click({ modifiers: ['Shift'] })
    await expect(ref).toHaveText('A1:E6')
    await cell(window, 4, 3).click({ modifiers: ['Shift'] })
    await expect(ref).toHaveText('A1:D5')
  })
})

test.describe('desfazer pelo menu fora da planilha', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('no documento, Editar > Desfazer e Refazer continuam no histórico do editor', async () => {
    const { window } = session
    await menu(session, 'new-document')
    const editor = window.locator('.ProseMirror')
    await editor.click()
    await window.keyboard.type('Presentes: Ana.')
    await expect(editor).toContainText('Presentes: Ana.')

    await clickMenuItem(session, 'Editar', 'Desfazer')
    await expect(editor).not.toContainText('Presentes')
    await clickMenuItem(session, 'Editar', 'Refazer')
    await expect(editor).toContainText('Presentes: Ana.')
  })
})
