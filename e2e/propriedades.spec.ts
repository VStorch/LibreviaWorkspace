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
test.describe('propriedades do documento', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-propriedades-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('salvar sem editar deixa as propriedades do arquivo como estavam', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithProperties())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    expect(await entryOf(origem, 'docProps/core.xml')).toBe(PROPERTIES_CORE)
    expect(await entryOf(origem, 'docProps/app.xml')).toBe(PROPERTIES_APP)
    expect(await entryOf(origem, 'docProps/custom.xml')).toBe(PROPERTIES_CUSTOM)
  })

  test('editar título e palavras-chave, salvar e reabrir mantém os valores', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithProperties())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'document-properties')
    const dialogo = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    const titulo = dialogo.getByRole('textbox', { name: 'Título' })
    await expect(titulo).toHaveValue('Relatório anual')
    await expect(titulo).toBeFocused()
    await expect(dialogo.getByRole('textbox', { name: 'Autor(es)' })).toHaveValue('Ana Lima')
    await expect(dialogo.getByRole('textbox', { name: 'Empresa' })).toHaveValue('ACME')
    // Read only: last saved by, revision, time and statistics.
    await expect(dialogo.getByRole('row', { name: /Modificado por/ })).toContainText('Bia')
    await expect(dialogo.getByRole('row', { name: /Revisão/ })).toContainText('7')
    await expect(dialogo.getByRole('row', { name: /Tempo total de edição/ })).toContainText('42 min')
    await expect(dialogo.getByRole('row', { name: /^Palavras/ })).toContainText('6')

    await titulo.fill('Relatório revisto')
    await dialogo.getByRole('textbox', { name: 'Palavras-chave' }).fill('contas; auditoria')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo).toHaveCount(0)
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Only `core.xml` changes: the edited fields and the save stamp. What the editor does not know
    // stays; `app.xml` and `custom.xml` go back byte for byte.
    const core = await entryOf(origem, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Relatório revisto</dc:title>')
    expect(core).toContain('<cp:keywords>contas; auditoria</cp:keywords>')
    expect(core).toContain('<cp:contentStatus>Rascunho</cp:contentStatus>')
    expect(core).toContain('<cp:revision>8</cp:revision>')
    expect(core).not.toContain('2025-02-03T04:05:06Z')
    expect(core).toContain('2025-01-02T03:04:05Z')
    expect(await entryOf(origem, 'docProps/app.xml')).toBe(PROPERTIES_APP)
    expect(await entryOf(origem, 'docProps/custom.xml')).toBe(PROPERTIES_CUSTOM)

    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')
    await menu(session, 'document-properties')
    await expect(dialogo.getByRole('textbox', { name: 'Título' })).toHaveValue('Relatório revisto')
    await expect(dialogo.getByRole('textbox', { name: 'Palavras-chave' })).toHaveValue('contas; auditoria')
    await expect(dialogo.getByRole('textbox', { name: 'Assunto' })).toHaveValue('Contas')
  })

  test('as personalizadas se editam, acrescentam e apagam, e só o custom.xml muda', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithProperties())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'document-properties')
    const dialogo = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    const nomes = dialogo.getByRole('textbox', { name: 'Nome' })
    const valores = dialogo.getByRole('textbox', { name: 'Valor' })
    await expect(nomes).toHaveValue('Cliente')
    await expect(valores).toHaveValue('XPTO')
    await valores.fill('ACME Ltda')

    await dialogo.getByRole('button', { name: 'Adicionar propriedade' }).click()
    await expect(nomes.nth(1)).toBeFocused()
    await nomes.nth(1).fill('Folhas')
    await dialogo.getByRole('combobox', { name: 'Tipo' }).nth(1).selectOption({ label: 'Número' })
    await dialogo.getByRole('spinbutton', { name: 'Valor' }).fill('3')

    // A row without a name does not go: the dialog says why and stays open.
    await dialogo.getByRole('button', { name: 'Adicionar propriedade' }).click()
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo.getByRole('alert')).toHaveText('Toda propriedade personalizada precisa de um nome.')
    await dialogo.getByRole('button', { name: 'Remover', exact: false }).last().click()
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await expect(dialogo).toHaveCount(0)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const custom = await entryOf(origem, 'docProps/custom.xml')
    expect(custom).toContain('pid="2" name="Cliente"><vt:lpwstr>ACME Ltda</vt:lpwstr>')
    expect(custom).toContain('pid="3" name="Folhas"><vt:i4>3</vt:i4>')
    expect(await entryOf(origem, 'docProps/app.xml')).toBe(PROPERTIES_APP)

    await menu(session, 'document-properties')
    await dialogo.getByRole('button', { name: 'Remover Cliente' }).click()
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(origem, 'docProps/custom.xml')).not.toContain('Cliente')

    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')
    await menu(session, 'document-properties')
    await expect(nomes).toHaveValue('Folhas')
    await expect(dialogo.getByRole('spinbutton', { name: 'Valor' })).toHaveValue('3')
  })
  test('o PDF exportado leva título, autor, assunto e palavras-chave', async () => {
    // Chromium writes the page `<title>` as the PDF Title; main appends the rest.
    const alvo = join(pasta, 'ata.pdf')
    await stubDialogs(session.app, { save: alvo, messageBox: 1 })
    await menu(session, 'new-document')
    await session.window.locator('.ProseMirror').click()
    await session.window.keyboard.type('Ata')

    await menu(session, 'document-properties')
    const dialogo = session.window.getByRole('dialog', { name: 'Propriedades do documento' })
    await dialogo.getByRole('textbox', { name: 'Título' }).fill('Ata de reuniao')
    await dialogo.getByRole('textbox', { name: 'Autor(es)' }).fill('Ana Lima')
    await dialogo.getByRole('textbox', { name: 'Assunto' }).fill('Conselho')
    await dialogo.getByRole('textbox', { name: 'Palavras-chave' }).fill('ata; reunião')
    await dialogo.getByRole('button', { name: 'Aplicar' }).click()

    await menu(session, 'export-pdf')
    const titulo = async (): Promise<string> => {
      const pdf = await readFile(alvo).catch(() => Buffer.alloc(0))
      return /\/Title \(([^)]*)\)/.exec(pdf.toString('latin1'))?.[1] ?? ''
    }
    // ASCII on purpose: outside it Chromium writes the title in UTF-16.
    await expect.poll(titulo, { timeout: 30_000 }).toBe('Ata de reuniao')

    if (await hasPdfinfo()) {
      const info = await pdfInfo(alvo)
      expect(info).toMatch(/Author:\s+Ana Lima/)
      expect(info).toMatch(/Subject:\s+Conselho/)
      expect(info).toMatch(/Keywords:\s+ata; reunião/)
    }
  })
})
