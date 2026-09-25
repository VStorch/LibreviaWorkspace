import { expect, test } from '@playwright/test'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithHeaderGrid, entryOf } from './fixtures.js'

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
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Título 1' })

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

  test('o .sdoc que veio de um .docx com cabeçalho volta a .docx avisando da faixa', async () => {
    // O caminho que o pacote mínimo abriu sem querer: o `.sdoc` guarda a faixa
    // com os endereços das relações do `.docx` de origem, e reaberto do disco
    // esse pacote não está mais aqui. A gravação parte do mínimo, que não tem
    // nenhuma dessas relações — e procurar por uma delas derrubava o sidecar
    // antes de gravar coisa alguma: nada no disco, "erro inesperado" na tela, e
    // a perda da faixa nunca chegava a ser dita.
    //
    // A faixa não tem onde ser gravada, e isso é perda inevitável. O que este
    // teste cobra é que o arquivo saia e que a perda apareça escrita.
    const origem = join(folder, 'grade.docx')
    const rascunho = join(folder, 'grade.sdoc')
    const destino = join(folder, 'volta.docx')
    const editor = session.window.locator('.ProseMirror')
    await writeFile(origem, await docxWithHeaderGrid())

    await stubDialogs(session.app, { open: origem, save: rascunho, messageBox: 1 })
    await menu(session, 'open')
    await expect(editor).toContainText('Primeira linha do corpo.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Reabrir do disco é o que apaga o original: o aplicativo guarda os bytes do
    // `.docx` na abertura, e quem abre o `.sdoc` não os tem.
    await stubDialogs(session.app, { open: rascunho })
    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()
    await menu(session, 'open')
    await expect(editor).toContainText('Primeira linha do corpo.')

    await stubDialogs(session.app, { save: destino })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Aviso de perda, e não faixa de erro: o `role="alert"` é do ErrorBanner.
    const aviso = session.window.locator('.banner--notice')
    await expect(aviso).toContainText('cabeçalho e rodapé do arquivo .docx de origem')
    await expect(session.window.getByRole('alert')).toHaveCount(0)

    // E o arquivo existe mesmo, com o corpo dentro.
    expect(await entryOf(destino, 'word/document.xml')).toContain('Primeira linha do corpo.')
  })
})
