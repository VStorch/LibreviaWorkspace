import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import {
  docxWithCellRevision,
  docxWithComment,
  docxWithFootnote,
  docxWithTrackedChange,
  docxWithoutExtras,
} from './fixtures.js'

/**
 * Graduated read-only: a structure revision (the inserted cell) locks, because editing the table
 * loses it. Comments, text revisions and notes do not lock: they go back to the file. Locking
 * everything would teach people to click "edit anyway" without reading.
 */
test.describe('read-only', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-ro-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('in English, the reason comes in English and still only once', async () => {
    const target = join(folder, 'minutes.docx')
    await writeFile(target, await docxWithCellRevision())
    await stubDialogs(session.app, { open: target, messageBox: 1 })
    await session.window.evaluate(async () => {
      const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } })
        .api
      await api.preferences.set({ language: 'en' })
    })

    await menu(session, 'open')

    await expect(session.window.locator('.banner--readonly')).toContainText('structure revisions')
    await expect(session.window.locator('.banner--notice')).toBeHidden()
  })

  test('a document with a structure revision opens locked and says why', async () => {
    const target = join(folder, 'ata.docx')
    await writeFile(target, await docxWithCellRevision())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    const banner = session.window.locator('.banner--readonly')
    await expect(banner).toBeVisible()
    await expect(banner).toContainText('revisões de estrutura')

    // A single warning: the inventory banner would repeat the reason.
    await expect(session.window.locator('.banner--notice')).toBeHidden()

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'false')

    // `setEditable` emits an update by default, which must not mark the document as modified.
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // The lock is a default, not a padlock: one click and editing comes back.
    await banner.getByRole('button', { name: 'Editar mesmo assim' }).click()
    await expect(banner).toBeHidden()
    await expect(editor).toHaveAttribute('contenteditable', 'true')

    await editor.click()
    await session.window.keyboard.type('Editado.')
    await expect(editor).toContainText('Editado.')
  })

  /**
   * The lock also applies to the menu: commands call the editor directly, and it obeys even with
   * `contenteditable="false"`.
   */
  test("the menu's edit commands respect the lock", async () => {
    const target = join(folder, 'ata-com-tabela.docx')
    await writeFile(target, await docxWithCellRevision({ leadingTable: true }))
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')
    await expect(session.window.locator('.banner--readonly')).toBeVisible()

    const lines = session.window.locator('.page__content tr')
    await expect(lines).toHaveCount(2)

    for (const command of [
      'table-row-after',
      'table-column-after',
      'table-merge-cells',
      'table-header-row',
      'table-delete',
      'insert-page-break',
      'paste-without-format',
    ]) {
      await menu(session, command)
    }

    // Commands with an editing dialog do not even open it.
    await menu(session, 'table-insert')
    await menu(session, 'table-properties')
    await menu(session, 'paragraph-setup')
    await menu(session, 'special-character')

    // Word count edits nothing, and still works.
    await menu(session, 'word-count')
    await expect(session.window.getByRole('dialog', { name: /Contagem de palavras/ })).toBeVisible()

    await expect(session.window.getByRole('dialog', { name: 'Inserir tabela' })).toHaveCount(0)
    await expect(session.window.getByRole('dialog', { name: 'Propriedades da tabela' })).toHaveCount(0)
    await expect(lines).toHaveCount(2)
    await expect(session.window.locator('.page__content table')).toHaveCount(1)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })

  test('a document with a comment opens editable and shows the comment', async () => {
    const target = join(folder, 'comentado.docx')
    await writeFile(target, await docxWithComment())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toBeHidden()
    await expect(session.window.locator('.banner--notice')).toBeHidden()
    await expect(session.window.locator('.comment-card')).toContainText('Conferir este número.')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })

  test('a document with track changes opens editable and shows the revisions', async () => {
    const target = join(folder, 'revisado.docx')
    await writeFile(target, await docxWithTrackedChange())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toBeHidden()
    await expect(session.window.locator('ins.revision').first()).toContainText('com uma inserção revisada.')
    await expect(session.window.locator('del.revision').first()).toContainText('Trecho excluído.')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })

  test('a document with a footnote opens editable and numbers the reference', async () => {
    const target = join(folder, 'nota.docx')
    await writeFile(target, await docxWithFootnote())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    await expect(session.window.locator('.pages__column .ProseMirror')).toHaveAttribute(
      'contenteditable',
      'true',
    )
    await expect(session.window.locator('.banner--readonly')).toBeHidden()
    const reference = session.window.locator('.page__content sup.note-ref')
    await expect(reference).toHaveAttribute('data-note-number', '1')
    await expect(session.window.locator('.paper-notes .note-body')).toContainText('Fonte: ata anterior.')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })

  test('a plain document opens editable', async () => {
    const target = join(folder, 'simples.docx')
    await writeFile(target, await docxWithoutExtras())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    await expect(session.window.locator('.pages__column .ProseMirror')).toHaveAttribute(
      'contenteditable',
      'true',
    )
    await expect(session.window.locator('.banner--readonly')).toBeHidden()
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })
})
