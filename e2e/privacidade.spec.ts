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
test.describe('inspetor de dados pessoais', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-privacidade-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  const dialogo = (): ReturnType<Session['window']['getByRole']> =>
    session.window.getByRole('dialog', { name: 'Inspecionar documento' })

  test('remove nomes e propriedades personalizadas, e o arquivo só muda ao salvar', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithProperties())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .ProseMirror')).toContainText('Relatório de contas')

    await menu(session, 'inspect-document')
    await expect(dialogo()).toContainText('Nomes de pessoas e da empresa')
    await expect(dialogo()).toContainText('Ana Lima, Bia, ACME')
    await expect(dialogo()).toContainText('Propriedades personalizadas')
    await expect(dialogo()).toContainText('Cliente')

    await dialogo().getByRole('button', { name: 'Remover' }).click()
    await expect(dialogo()).toContainText('Removido.')
    expect(await entryOf(origem, 'docProps/core.xml')).toBe(PROPERTIES_CORE)
    await dialogo().getByRole('button', { name: 'Fechar' }).click()
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const core = await entryOf(origem, 'docProps/core.xml')
    expect(core).toContain('<dc:title>Relatório anual</dc:title>')
    expect(core).not.toContain('Ana Lima')
    expect(core).not.toContain('lastModifiedBy')
    expect(await entryOf(origem, 'docProps/app.xml')).not.toContain('ACME')
    await expect(entryOf(origem, 'docProps/custom.xml')).rejects.toThrow()
    expect(await entryOf(origem, 'word/settings.xml')).toContain('<w:removePersonalInformation')

    await menu(session, 'inspect-document')
    await expect(dialogo()).toContainText('Nada encontrado')
    await expect(dialogo()).toContainText('já pede para não guardar nomes')
  })

  test('os autores dos comentários passam a "Autor", também nos comentários novos', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithComment())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes.first()).toContainText('Revisor')

    await menu(session, 'inspect-document')
    await expect(dialogo()).toContainText('Revisor')
    await dialogo().getByRole('button', { name: 'Remover' }).click()
    await expect(dialogo()).toContainText('Removido.')
    await dialogo().getByRole('button', { name: 'Fechar' }).click()
    await expect(cartoes.first()).toContainText('Autor')
    await expect(cartoes.first()).not.toContainText('Revisor')

    await session.window.locator('.ProseMirror').getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    await session.window.keyboard.type('Sem nome.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes.filter({ hasText: 'Sem nome.' })).toContainText('Autor')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const comentarios = await entryOf(origem, 'word/comments.xml')
    expect(comentarios).not.toContain('Revisor')
    expect(comentarios.match(/w:author="Autor"/g)).toHaveLength(2)
  })

  test('com alterações não salvas, salva antes de inspecionar', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithProperties())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await session.window.locator('.ProseMirror').getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Assinado.')

    await menu(session, 'inspect-document')
    await expect(dialogo()).toContainText('Salve antes')
    await dialogo().getByRole('button', { name: 'Salvar e inspecionar' }).click()

    await expect(dialogo()).toContainText('Ana Lima')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(origem, 'word/document.xml')).toContain('Assinado.')
  })
})
