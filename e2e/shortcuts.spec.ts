import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

async function acceleratorOf(session: Session, top: string, label: string): Promise<string | null> {
  return session.app.evaluate(
    ({ Menu }, [topLabel, itemLabel]) => {
      const submenu = Menu.getApplicationMenu()?.items.find((entry) => entry.label === topLabel)?.submenu
      const item = submenu?.items.find((entry) => entry.label === itemLabel)
      return (item?.accelerator as string | undefined) ?? null
    },
    [top, label] as const,
  )
}

test.describe('custom shortcuts', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('a new key moves the menu accelerator and the toolbar hint, and Restore all brings the default back', async () => {
    const { window } = session
    await menu(session, 'new-document')
    await expect(window.locator('.ProseMirror')).toBeFocused()
    await menu(session, 'options')
    await window
      .getByRole('dialog', { name: 'Opções' })
      .getByRole('button', { name: 'Personalizar atalhos…' })
      .click()
    const dialog = window.getByRole('dialog', { name: 'Atalhos do teclado' })

    await dialog.getByRole('button', { name: 'Imprimir…: Ctrl+P' }).click()
    await window.keyboard.press('Control+b')
    await expect(dialog.getByRole('alert')).toHaveText(
      'Ctrl+B já é do editor de texto, como negrito ou alinhamento.',
    )
    await window.keyboard.press('Control+o')
    await expect(dialog.getByRole('alert')).toHaveText('Ctrl+O já é de “Abrir…”.')
    await window.keyboard.press('Escape')
    await expect(dialog.getByRole('button', { name: 'Imprimir…: Ctrl+P' })).toBeVisible()

    await dialog.getByRole('button', { name: 'Localizar e substituir…: Ctrl+F' }).click()
    await window.keyboard.press('Control+Shift+k')
    await expect(dialog.getByRole('button', { name: 'Localizar e substituir…: Ctrl+Shift+K' })).toBeVisible()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(window.getByRole('dialog', { name: 'Opções' })).toBeVisible()
    await window.getByRole('dialog', { name: 'Opções' }).getByRole('button', { name: 'Cancelar' }).click()

    await expect
      .poll(() => acceleratorOf(session, 'Editar', 'Localizar e substituir…'))
      .toBe('CmdOrCtrl+Shift+K')
    await expect(window.getByRole('button', { name: 'Localizar e substituir' })).toHaveAttribute(
      'title',
      'Localizar e substituir (Ctrl+Shift+K)',
    )

    await menu(session, 'options')
    await window
      .getByRole('dialog', { name: 'Opções' })
      .getByRole('button', { name: 'Personalizar atalhos…' })
      .click()
    await dialog.getByRole('button', { name: 'Restaurar todos' }).click()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect.poll(() => acceleratorOf(session, 'Editar', 'Localizar e substituir…')).toBe('CmdOrCtrl+F')
  })
})
