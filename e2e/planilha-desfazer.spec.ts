import { expect, test } from '@playwright/test'
import { clickMenuItem, launch, menu, type Session } from './app.js'
import { cell, gridReady, reference, select, selectDown, write } from './sheet-grid.js'

test.describe('undo in the spreadsheet', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-spreadsheet')
    await gridReady(session.window)
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('Ctrl+Z undoes, Ctrl+Y and Ctrl+Shift+Z redo, and undoing records no new step', async () => {
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

  test('after undoing, the active cell shows the earlier value', async () => {
    const { window } = session
    await write(window, 0, 1, 'b')
    await select(window, 0, 1)
    await window.keyboard.press('Control+z')
    await expect(cell(window, 0, 1)).toHaveText('')
    await expect(window.locator('.formula-bar__input')).toHaveValue('')
  })

  test('undoing back to what was saved clears the modified notice', async () => {
    const { window } = session
    const title = await window.title()
    await write(window, 0, 0, 'a')
    await expect(window).not.toHaveTitle(title)
    await window.keyboard.press('Control+z')
    await expect(window).toHaveTitle(title)
  })

  test('Edit > Undo and Redo in the menu use the spreadsheet history', async () => {
    const { window } = session
    await write(window, 0, 0, 'a')
    await cell(window, 5, 3).click()
    await clickMenuItem(session, 'Editar', 'Desfazer')
    await expect(cell(window, 0, 0)).toHaveText('')
    await clickMenuItem(session, 'Editar', 'Refazer')
    await expect(cell(window, 0, 0)).toHaveText('a')
  })

  test('with a dialog open, Ctrl+Z does not undo the spreadsheet behind it', async () => {
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
  test('Ctrl+Shift+Y is not redo: the spreadsheet leaves the key for the assistant shortcut', async () => {
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

  test('text starting with "=" stays text when edited again', async () => {
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

test.describe('spreadsheet selection', () => {
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

  test('right-click inside the selection keeps the selection and the menu acts on all of it', async () => {
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

  test('right-click outside the selection goes to the clicked cell', async () => {
    const { window } = session
    await selectBlock()
    await cell(window, 5, 3).click({ button: 'right' })
    await expect(window.locator('.formula-bar__ref')).toHaveText('D6')
    await expect(window.getByRole('menuitem', { name: 'Excluir linha', exact: true })).toBeVisible()
  })

  test('with frozen columns, the selected cell and range have the right reference', async () => {
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

  test('with frozen columns, Shift+arrow and Shift+click extend the selection across the divider', async () => {
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

test.describe('undo from the menu outside the spreadsheet', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('in the document, Edit > Undo and Redo stay on the editor history', async () => {
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
