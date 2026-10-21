import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { PROPERTIES_CORE, docxWithComment, docxWithProperties, entryOf } from './fixtures.js'

/**
 * File → Inspect document: names, custom properties and authors leave the file only when saving,
 * and the file keeps asking not to sign.
 */
test.describe('personal data inspector', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-privacy-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  const dialog = (): ReturnType<Session['window']['getByRole']> =>
    session.window.getByRole('dialog', { name: 'Inspecionar documento' })

  test('removes names and custom properties, and the file only changes on save', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'inspect-document')
    await expect(dialog()).toContainText('Nomes de pessoas e da empresa')
    await expect(dialog()).toContainText('Ana Lima, Bia, ACME')
    await expect(dialog()).toContainText('Propriedades personalizadas')
    await expect(dialog()).toContainText('Cliente')

    await dialog().getByRole('button', { name: 'Remover' }).click()
    await expect(dialog()).toContainText('Removido.')
    expect(await entryOf(sourceFile, 'docProps/core.xml')).toBe(PROPERTIES_CORE)
    await dialog().getByRole('button', { name: 'Fechar' }).click()
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const core = await entryOf(sourceFile, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Relatório anual</dc:title>')
    expect(core).not.toContain('Ana Lima')
    expect(core).not.toContain('lastModifiedBy')
    expect(await entryOf(sourceFile, 'docProps/app.xml')).not.toContain('ACME')
    await expect(entryOf(sourceFile, 'docProps/custom.xml')).rejects.toThrow()
    expect(await entryOf(sourceFile, 'word/settings.xml')).toContain('<w:removePersonalInformation')

    await menu(session, 'inspect-document')
    await expect(dialog()).toContainText('Nada encontrado')
    await expect(dialog()).toContainText('já pede para não guardar nomes')
  })

  test('comment authors become "Autor", also for new comments', async () => {
    const sourceFile = join(directory, 'ata.docx')
    await writeFile(sourceFile, await docxWithComment())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    const cards = session.window.locator('.comment-card')
    await expect(cards.first()).toContainText('Revisor')

    await menu(session, 'inspect-document')
    await expect(dialog()).toContainText('Revisor')
    await dialog().getByRole('button', { name: 'Remover' }).click()
    await expect(dialog()).toContainText('Removido.')
    await dialog().getByRole('button', { name: 'Fechar' }).click()
    await expect(cards.first()).toContainText('Autor')
    await expect(cards.first()).not.toContainText('Revisor')

    await session.window.locator('.ProseMirror').getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    await session.window.keyboard.type('Sem nome.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cards.filter({ hasText: 'Sem nome.' })).toContainText('Autor')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const comments = await entryOf(sourceFile, 'word/comments.xml')
    expect(comments).not.toContain('Revisor')
    expect(comments.match(/w:author="Autor"/g)).toHaveLength(2)
  })

  test('with unsaved changes, saves before inspecting', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithProperties())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await session.window.locator('.ProseMirror').getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Assinado.')

    await menu(session, 'inspect-document')
    await expect(dialog()).toContainText('Salve antes')
    await dialog().getByRole('button', { name: 'Salvar e inspecionar' }).click()

    await expect(dialog()).toContainText('Ana Lima')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(sourceFile, 'word/document.xml')).toContain('Assinado.')
  })
})
