import { mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { clickMenuItem, launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, entryOf } from './fixtures.js'

const UNTOUCHED =
  '<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"><w:r><w:t>Parágrafo intocado.</w:t></w:r></w:p>'
const paragraph = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

test.describe('tabs', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-tabs-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  async function open(path: string): Promise<void> {
    await stubDialogs(session.app, { open: path })
    await menu(session, 'open')
  }

  async function twoFiles(first: string): Promise<{ a: string; b: string }> {
    const a = join(folder, 'a.docx')
    const b = join(folder, 'b.docx')
    await writeFile(a, await docxWithNamedStyles(first))
    await writeFile(b, await docxWithNamedStyles(paragraph('Texto de B.')))
    return { a, b }
  }

  const tabs = () => session.window.locator('.tabbar__tab')
  const editor = () => session.window.locator('.ProseMirror')
  const tabNamed = (name: string) => session.window.locator('.tabbar__label', { hasText: name })

  test('each file opens in a tab of its own, and each tab shows its own content', async () => {
    const { a, b } = await twoFiles(paragraph('Texto de A.'))

    await open(a)
    await expect(editor()).toContainText('Texto de A.')
    await open(b)
    await expect(editor()).toContainText('Texto de B.')
    // Home gave its place to the first file.
    await expect(tabs()).toHaveCount(2)

    await tabNamed('a.docx').click()
    await expect(editor()).toContainText('Texto de A.')
    await expect(editor()).not.toContainText('Texto de B.')

    await open(b)
    await expect(tabs()).toHaveCount(2)
    await expect(editor()).toContainText('Texto de B.')
  })

  test('picking two files in the open dialog opens both, each in its tab', async () => {
    const { a, b } = await twoFiles(paragraph('Texto de A.'))

    await stubDialogs(session.app, { open: [a, b] })
    await menu(session, 'open')

    await expect(tabs()).toHaveCount(2)
    await expect(editor()).toContainText('Texto de B.')
    await tabNamed('a.docx').click()
    await expect(editor()).toContainText('Texto de A.')
  })

  test('saving A after opening B keeps what A did not edit byte for byte', async () => {
    const { a, b } = await twoFiles(UNTOUCHED + paragraph('Parágrafo a editar.'))

    await open(a)
    await expect(editor()).toContainText('Parágrafo a editar.')
    await open(b)
    await expect(editor()).toContainText('Texto de B.')

    await tabNamed('a.docx').click()
    await editor().getByText('Parágrafo a editar.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Editado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const original = await docxWithNamedStyles(UNTOUCHED)
    const sourceCopy = join(folder, 'source.docx')
    await writeFile(sourceCopy, original)
    expect(await entryOf(a, 'word/styles.xml')).toBe(await entryOf(sourceCopy, 'word/styles.xml'))
    const body = await entryOf(a, 'word/document.xml')
    expect(body).toContain(UNTOUCHED)
    expect(body).toContain('Editado.')
  })

  test('closing a modified tab asks first, and closing the last leaves Home', async () => {
    await menu(session, 'new-document')
    await editor().click()
    await session.window.keyboard.type('Texto sem salvar')

    await stubDialogs(session.app, { messageBox: 2 })
    await session.window.locator('.tabbar__tab--active .tabbar__close').click()
    await expect(editor()).toContainText('Texto sem salvar')

    await stubDialogs(session.app, { messageBox: 1 })
    await session.window.locator('.tabbar__tab--active .tabbar__close').click()
    await expect(session.window.locator('.home')).toBeVisible()
    await expect(tabs()).toHaveCount(1)
    await expect(tabs()).toContainText('Início')
  })

  test('a hidden tab autosaves what was typed in it, and the draft comes back', async () => {
    await menu(session, 'new-document')
    await editor().click()
    await session.window.keyboard.type('Texto da aba oculta')
    // A second tab hides the first before any autosave has run.
    await menu(session, 'new-document')
    await expect(tabs()).toHaveCount(2)
    // The dot in the title follows the tab on screen, which has nothing unsaved.
    await expect
      .poll(() => session.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]!.getTitle()))
      .not.toMatch(/^•/)

    await expect
      .poll(async () => draftsOnDisk(session), { timeout: 20_000, message: 'the draft was not written' })
      .toContain('Texto da aba oculta')

    const { userData } = session
    await session.crash()
    session = await launch({ userData })

    await session.window.locator('.banner--recovery').getByRole('button', { name: 'Recuperar' }).click()
    await expect(editor()).toContainText('Texto da aba oculta')
  })

  test('after a crash with two unsaved tabs, Recover all brings both back', async () => {
    for (const text of ['Primeira aba', 'Segunda aba']) {
      await menu(session, 'new-document')
      await editor().click()
      await session.window.keyboard.type(text)
    }
    await expect
      .poll(async () => draftsOnDisk(session), { timeout: 20_000, message: 'the drafts were not written' })
      .toMatch(/Primeira aba[\s\S]*Segunda aba|Segunda aba[\s\S]*Primeira aba/)

    const { userData } = session
    await session.crash()
    session = await launch({ userData })

    const banner = session.window.locator('.banner--recovery')
    await expect(banner).toContainText('Há mais 1 rascunho.')
    await banner.getByRole('button', { name: 'Recuperar todos' }).click()
    await expect(tabs()).toHaveCount(2)
    await expect(banner).toBeHidden()
  })

  const activeTab = () => session.window.locator('.tabbar__tab--active .tabbar__label')

  test('the next and previous tab commands move between tabs', async () => {
    const { a, b } = await twoFiles(paragraph('Texto de A.'))
    await open(a)
    await open(b)

    // Native accelerators do not hear synthetic keys: Ctrl+Tab is the menu command.
    await menu(session, 'next-tab')
    await expect(activeTab()).toHaveAccessibleName('a.docx')
    await menu(session, 'previous-tab')
    await expect(activeTab()).toHaveAccessibleName('b.docx')
  })

  test('on a spreadsheet Ctrl+PgDn walks forward through the sheets, as Ctrl+PgUp walks back', async () => {
    await menu(session, 'new-spreadsheet')
    await session.window.locator('.tabs__add').click()
    await session.window.locator('.tabs__add').click()
    const sheets = session.window.locator('.workbook [role="tab"]')
    await expect(sheets).toHaveCount(3)
    await sheets.first().click()
    await session.window.locator('revo-grid').click()

    await session.window.keyboard.press('Control+PageDown')
    await expect(sheets.nth(1)).toHaveAttribute('aria-selected', 'true')
    await session.window.keyboard.press('Control+PageDown')
    await expect(sheets.nth(2)).toHaveAttribute('aria-selected', 'true')
    await session.window.keyboard.press('Control+PageUp')
    await expect(sheets.nth(1)).toHaveAttribute('aria-selected', 'true')
  })

  test('closing all tabs asks about a modified one, and keeping it stops there', async () => {
    const { a } = await twoFiles(paragraph('Texto de A.'))
    await open(a)
    await menu(session, 'new-document')
    await editor().click()
    await session.window.keyboard.type('Texto sem salvar')

    await stubDialogs(session.app, { messageBox: 2 })
    await menu(session, 'close-all-tabs')
    await expect(tabs()).toHaveCount(1)
    await expect(editor()).toContainText('Texto sem salvar')

    await stubDialogs(session.app, { messageBox: 1 })
    await menu(session, 'close-all-tabs')
    await expect(session.window.locator('.home')).toBeVisible()
    await expect(tabs()).toHaveCount(1)
  })

  test('File → Home shows the Home tab beside the open file', async () => {
    const { a } = await twoFiles(paragraph('Texto de A.'))
    await open(a)

    await clickMenuItem(session, 'Arquivo', 'Início')

    await expect(session.window.locator('.home')).toBeVisible()
    await expect(tabs()).toHaveCount(2)
  })

  test('the track changes check mark follows the tab on screen', async () => {
    const checked = () =>
      session.app.evaluate(({ Menu }) => {
        const review = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Revisão')
        return review?.submenu?.items[0]?.checked ?? false
      })

    await menu(session, 'new-document')
    await editor().click()
    await menu(session, 'toggle-track-changes')
    await expect.poll(checked).toBe(true)

    await menu(session, 'new-document')
    await expect(tabs()).toHaveCount(2)
    await expect.poll(checked).toBe(false)

    await menu(session, 'previous-tab')
    await expect.poll(checked).toBe(true)
  })
  test('the menu greys out what the tab on screen cannot do', async () => {
    const enabled = (top: string, label: string) =>
      session.app.evaluate(
        ({ Menu }, [menuLabel, itemLabel]) => {
          const submenu = Menu.getApplicationMenu()?.items.find((entry) => entry.label === menuLabel)?.submenu
          return submenu?.items.find((entry) => entry.label === itemLabel)?.enabled
        },
        [top, label] as const,
      )

    await expect.poll(() => enabled('Arquivo', 'Salvar')).toBe(false)
    await expect.poll(() => enabled('Inserir', 'Comentário')).toBe(false)

    await menu(session, 'new-document')
    await expect.poll(() => enabled('Arquivo', 'Salvar')).toBe(true)
    await expect.poll(() => enabled('Inserir', 'Comentário')).toBe(true)

    await menu(session, 'new-spreadsheet')
    await expect.poll(() => enabled('Inserir', 'Comentário')).toBe(false)
    await expect.poll(() => enabled('Arquivo', 'Salvar')).toBe(true)
  })

  async function twoModifiedTabs(): Promise<void> {
    const { a, b } = await twoFiles(paragraph('Texto de A.'))
    for (const path of [a, b]) {
      await open(path)
      await editor().click()
      await session.window.keyboard.type('Editado. ')
    }
    // Main hears the tab list with the title: once the dot is there, it knows both.
    await expect
      .poll(() => session.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]!.getTitle()))
      .toMatch(/^•/)
  }

  const closeWindow = () =>
    session.app.evaluate(({ BrowserWindow }) => {
      BrowserWindow.getAllWindows()[0]?.close()
    })

  test('quitting with two modified tabs asks once, naming both; Cancel keeps them', async () => {
    await twoModifiedTabs()
    await session.app.evaluate(({ dialog }) => {
      const asked: string[] = []
      Object.assign(globalThis, { asked })
      dialog.showMessageBox = (async (...args: unknown[]) => {
        const options = args.at(-1) as Electron.MessageBoxOptions
        asked.push(`${options.message}\n${options.detail ?? ''}\n${options.buttons?.join('|') ?? ''}`)
        return { response: 2, checkboxChecked: false }
      }) as typeof dialog.showMessageBox
    })

    await closeWindow()

    const asked = () => session.app.evaluate(() => (globalThis as unknown as { asked: string[] }).asked)
    await expect.poll(async () => (await asked()).length).toBe(1)
    const [question] = await asked()
    expect(question).toContain('2 arquivos')
    expect(question).toContain('a.docx\nb.docx')
    expect(question).toContain('Salvar todos|Não salvar|Cancelar')
    await expect(tabs()).toHaveCount(2)
    await expect(editor()).toContainText('Editado.')
  })

  test('quitting with two modified tabs and choosing Don’t save closes the window', async () => {
    await twoModifiedTabs()
    await stubDialogs(session.app, { messageBox: 1 })
    const closed = session.app.waitForEvent('close')

    await closeWindow()

    await closed
    await rm(session.userData, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 })
    session = await launch()
  })
})

async function draftsOnDisk(session: Session): Promise<string> {
  const folder = join(session.userData, 'recovery', 'drafts')
  const names = await readdir(folder).catch(() => [])
  const contents = await Promise.all(names.map((name) => readFile(join(folder, name), 'utf8')))
  return contents.join('\n')
}
