import { expect, test } from '@playwright/test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { entryOf } from './fixtures.js'

/**
 * O documento que nasceu no editor, salvo como `.docx`.
 *
 * Antes era recusado: a gravação cirúrgica pressupõe um original, e o documento
 * novo não tem. Agora o sidecar cria um pacote mínimo que faz esse papel, e toda
 * gravação parte dele — inclusive a segunda, para que o arquivo não some uma
 * camada por Ctrl+S. O teste percorre o caminho inteiro — diálogo, sidecar
 * publicado, disco — porque é nas fronteiras entre eles que esse tipo de ligação
 * se perde.
 */
test.describe('documento novo em .docx', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-docx-novo-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('salva como .docx, grava de novo no mesmo arquivo e reabre com título e texto', async () => {
    const destino = join(folder, 'relatorio.docx')
    const editor = session.window.locator('.ProseMirror')

    await menu(session, 'new-document')
    await editor.click()
    await session.window.keyboard.type('Relatório anual')
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption('1')

    await editor.locator('h1').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Primeiro parágrafo.')

    await stubDialogs(session.app, { save: destino, open: destino })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // O título aponta um estilo que o pacote define — e não um nome solto.
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toContain('<w:pStyle w:val="Heading1"')
    expect(corpo).toContain('Primeiro parágrafo.')
    expect(await entryOf(destino, 'word/styles.xml')).toContain('w:styleId="Heading1"')

    // A segunda gravação vai ao mesmo arquivo sem perguntar nada: o caminho
    // ficou autorizado, e ela parte do mesmo pacote mínimo da primeira.
    await session.window.keyboard.type(' Segunda gravação.')
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(destino, 'word/document.xml')).toContain('Segunda gravação.')

    await menu(session, 'open')
    await expect(editor.locator('h1')).toHaveText('Relatório anual')
    await expect(editor.locator('p').first()).toHaveText('Primeiro parágrafo. Segunda gravação.')
  })
})
