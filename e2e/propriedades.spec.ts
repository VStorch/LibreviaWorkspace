import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { hasPdfinfo, pdfInfo } from './external-tools.js'
import {
  PROPERTIES_APP,
  PROPERTIES_CORE,
  PROPERTIES_CUSTOM,
  docxWithProperties,
  entryOf,
} from './fixtures.js'

/**
 * File → Properties: title, subject, author… read from `docProps/`, edited in the dialog and
 * written only to the part that changed.
 */
test.describe('document properties', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-properties-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('saving without editing leaves the file properties as they were', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    expect(await entryOf(sourceFile, 'docProps/core.xml')).toBe(PROPERTIES_CORE)
    expect(await entryOf(sourceFile, 'docProps/app.xml')).toBe(PROPERTIES_APP)
    expect(await entryOf(sourceFile, 'docProps/custom.xml')).toBe(PROPERTIES_CUSTOM)
  })

  test('editing title and keywords, saving and reopening keeps the values', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'document-properties')
    const dialog = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    const title = dialog.getByRole('textbox', { name: 'Título' })
    await expect(title).toHaveValue('Relatório anual')
    await expect(title).toBeFocused()
    await expect(dialog.getByRole('textbox', { name: 'Autor(es)' })).toHaveValue('Ana Lima')
    await expect(dialog.getByRole('textbox', { name: 'Empresa' })).toHaveValue('ACME')
    // Read only: last saved by, revision, time and statistics.
    await expect(dialog.getByRole('row', { name: /Modificado por/ })).toContainText('Bia')
    await expect(dialog.getByRole('row', { name: /Revisão/ })).toContainText('7')
    await expect(dialog.getByRole('row', { name: /Tempo total de edição/ })).toContainText('42 min')
    await expect(dialog.getByRole('row', { name: /^Palavras/ })).toContainText('6')

    await title.fill('Relatório revisto')
    await dialog.getByRole('textbox', { name: 'Palavras-chave' }).fill('contas; auditoria')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialog).toHaveCount(0)
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Only `core.xml` changes: the edited fields and the save stamp. What the editor does not know
    // stays; `app.xml` and `custom.xml` go back byte for byte.
    const core = await entryOf(sourceFile, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Relatório revisto</dc:title>')
    expect(core).toContain('<cp:keywords>contas; auditoria</cp:keywords>')
    expect(core).toContain('<cp:contentStatus>Rascunho</cp:contentStatus>')
    expect(core).toContain('<cp:revision>8</cp:revision>')
    expect(core).not.toContain('2025-02-03T04:05:06Z')
    expect(core).toContain('2025-01-02T03:04:05Z')
    expect(await entryOf(sourceFile, 'docProps/app.xml')).toBe(PROPERTIES_APP)
    expect(await entryOf(sourceFile, 'docProps/custom.xml')).toBe(PROPERTIES_CUSTOM)

    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')
    await menu(session, 'document-properties')
    await expect(dialog.getByRole('textbox', { name: 'Título' })).toHaveValue('Relatório revisto')
    await expect(dialog.getByRole('textbox', { name: 'Palavras-chave' })).toHaveValue('contas; auditoria')
    await expect(dialog.getByRole('textbox', { name: 'Assunto' })).toHaveValue('Contas')
  })

  test('custom ones can be edited, added and deleted, and only custom.xml changes', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'document-properties')
    const dialog = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    const names = dialog.getByRole('textbox', { name: 'Nome' })
    const values = dialog.getByRole('textbox', { name: 'Valor' })
    await expect(names).toHaveValue('Cliente')
    await expect(values).toHaveValue('XPTO')
    await values.fill('ACME Ltda')

    await dialog.getByRole('button', { name: 'Adicionar propriedade' }).click()
    await expect(names.nth(1)).toBeFocused()
    await names.nth(1).fill('Folhas')
    await dialog.getByRole('combobox', { name: 'Tipo' }).nth(1).selectOption({ label: 'Número' })
    await dialog.getByRole('spinbutton', { name: 'Valor' }).fill('3')

    // A row without a name does not go: the dialog says why and stays open.
    await dialog.getByRole('button', { name: 'Adicionar propriedade' }).click()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialog.getByRole('alert')).toHaveText('Toda propriedade personalizada precisa de um nome.')
    await dialog.getByRole('button', { name: 'Remover', exact: false }).last().click()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialog).toHaveCount(0)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const custom = await entryOf(sourceFile, 'docProps/custom.xml')
    expect(custom).toContain('pid="2" name="Cliente"><vt:lpwstr>ACME Ltda</vt:lpwstr>')
    expect(custom).toContain('pid="3" name="Folhas"><vt:i4>3</vt:i4>')
    expect(await entryOf(sourceFile, 'docProps/app.xml')).toBe(PROPERTIES_APP)

    await menu(session, 'document-properties')
    await dialog.getByRole('button', { name: 'Remover Cliente' }).click()
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(sourceFile, 'docProps/custom.xml')).not.toContain('Cliente')

    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')
    await menu(session, 'document-properties')
    await expect(names).toHaveValue('Folhas')
    await expect(dialog.getByRole('spinbutton', { name: 'Valor' })).toHaveValue('3')
  })
  test('status goes to core and, with edited text, statistics go to app.xml', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    const text = session.window.locator('.pages__column .ProseMirror')
    await expect(text).toContainText('Relatório de contas')

    await menu(session, 'document-properties')
    const dialog = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    const status = dialog.getByRole('textbox', { name: 'Status' })
    await expect(status).toHaveValue('Rascunho')
    await status.fill('Final')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(sourceFile, 'docProps/core.xml')).toContain(
      '<cp:contentStatus>Final</cp:contentStatus>',
    )
    expect(await entryOf(sourceFile, 'docProps/app.xml')).toBe(PROPERTIES_APP)

    await text.click()
    await session.window.keyboard.press('Control+End')
    await session.window.keyboard.type(' Mais duas.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const app = await entryOf(sourceFile, 'docProps/app.xml')
    expect(app).toContain('<Pages>1</Pages>')
    expect(app).toMatch(/<Words>\d+<\/Words><Characters>\d+<\/Characters>/)
    expect(app).toContain('<Template>Normal.dotm</Template>')
  })

  test('the property field shows the value and follows changes to the properties', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    const text = session.window.locator('.pages__column .ProseMirror')
    await expect(text).toContainText('Relatório de contas')
    await text.click()
    await session.window.keyboard.press('Control+End')
    await session.window.keyboard.type(' Cliente: ')

    await menu(session, 'insert-property-field')
    const insert = session.window.getByRole('dialog', { name: 'Inserir propriedade' })
    await insert.getByRole('listbox', { name: 'Propriedade' }).selectOption({ label: 'Cliente' })
    await expect(insert).toContainText('Valor atual: XPTO')
    await insert.getByRole('button', { name: 'Inserir' }).click()
    const field = text.locator('[data-field="docproperty"]')
    await expect(field).toHaveText('XPTO')

    await menu(session, 'document-properties')
    const dialog = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    await dialog.getByRole('textbox', { name: 'Valor' }).fill('ACME Ltda')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(field).toHaveText('ACME Ltda')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const document = await entryOf(sourceFile, 'word/document.xml')
    expect(document).toContain('DOCPROPERTY "Cliente" \\* MERGEFORMAT')
    expect(document).toContain('ACME Ltda')
  })

  test('the exported PDF carries title, author, subject and keywords', async () => {
    // Chromium writes the page `<title>` as the PDF Title; main appends the rest.
    const targetPoint = join(directory, 'ata.pdf')
    await stubDialogs(session.app, { save: targetPoint, messageBox: 1 })
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Ata')

    await menu(session, 'document-properties')
    const dialog = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    await dialog.getByRole('textbox', { name: 'Título' }).fill('Ata de reuniao')
    await dialog.getByRole('textbox', { name: 'Autor(es)' }).fill('Ana Lima')
    await dialog.getByRole('textbox', { name: 'Assunto' }).fill('Conselho')
    await dialog.getByRole('textbox', { name: 'Palavras-chave' }).fill('ata; reunião')
    await dialog.getByRole('button', { name: 'Aplicar' }).click()

    await menu(session, 'export-pdf')
    const title = async (): Promise<string> => {
      const pdf = await readFile(targetPoint).catch(() => Buffer.alloc(0))
      return /\/Title \(([^)]*)\)/.exec(pdf.toString('latin1'))?.[1] ?? ''
    }
    // ASCII on purpose: outside it Chromium writes the title in UTF-16.
    await expect.poll(title, { timeout: 30_000 }).toBe('Ata de reuniao')

    if (await hasPdfinfo()) {
      const info = await pdfInfo(targetPoint)
      expect(info).toMatch(/Author:\s+Ana Lima/)
      expect(info).toMatch(/Subject:\s+Conselho/)
      expect(info).toMatch(/Keywords:\s+ata; reunião/)
    }
  })
})
