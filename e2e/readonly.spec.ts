import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithComment, docxWithoutExtras } from './fixtures.js'

/**
 * Somente leitura **graduado**, não ligado/desligado.
 *
 * Documento com comentário abre travado, porque editar o parágrafo que ancora o
 * comentário o apaga. Documento comum abre editável, porque travar tudo
 * ensinaria o usuário a clicar "editar mesmo assim" sem ler — e aí a proteção
 * deixaria de proteger.
 */
test.describe('somente leitura', () => {
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

  test('documento com comentário abre travado e diz por quê', async () => {
    const target = join(folder, 'ata.docx')
    await writeFile(target, await docxWithComment())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    const banner = session.window.locator('.banner--readonly')
    await expect(banner).toBeVisible()
    await expect(banner).toContainText('comentários')

    // Um aviso só: a faixa de inventário repetiria "comentários" logo abaixo, e
    // dois avisos dizendo a mesma coisa valem menos que um.
    await expect(session.window.locator('.banner--notice')).toBeHidden()

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'false')

    // Travar o editor não pode marcar o documento como alterado: `setEditable`
    // emite um update por padrão, e com ele todo arquivo aberto apareceria como
    // "não salvo" antes de o usuário tocar em nada.
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // A trava é um padrão, não um cadeado: um clique e a edição volta.
    await banner.getByRole('button', { name: 'Editar mesmo assim' }).click()
    await expect(banner).toBeHidden()
    await expect(editor).toHaveAttribute('contenteditable', 'true')

    await editor.click()
    await session.window.keyboard.type('Editado.')
    await expect(editor).toContainText('Editado.')
  })

  /**
   * A trava vale para o que chega pelo menu, e não só para o teclado.
   *
   * O `contenteditable="false"` segura a digitação, mas os comandos do menu
   * chamam o editor direto — e o editor obedece a comando mesmo travado. Inserir
   * linha, excluir a tabela, sombrear pelas propriedades e inserir quebra de
   * página funcionavam num documento aberto em somente leitura, e o status
   * virava "Não salvo".
   */
  test('os comandos de edição do menu respeitam a trava', async () => {
    const target = join(folder, 'ata-com-tabela.docx')
    await writeFile(target, await docxWithComment({ leadingTable: true }))
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')
    await expect(session.window.locator('.banner--readonly')).toBeVisible()

    const linhas = session.window.locator('.page__content tr')
    await expect(linhas).toHaveCount(2)

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

    // Os comandos que abrem diálogo de edição nem abrem: um diálogo que aceita e
    // não aplica seria pior que nenhum.
    await menu(session, 'table-insert')
    await menu(session, 'table-properties')
    await menu(session, 'paragraph-setup')
    await menu(session, 'special-character')

    // A contagem de palavras não edita nada, e continua valendo.
    await menu(session, 'word-count')
    await expect(session.window.getByRole('dialog', { name: /Contagem de palavras/ })).toBeVisible()

    await expect(session.window.getByRole('dialog', { name: 'Inserir tabela' })).toHaveCount(0)
    await expect(session.window.getByRole('dialog', { name: 'Propriedades da tabela' })).toHaveCount(0)
    await expect(linhas).toHaveCount(2)
    await expect(session.window.locator('.page__content table')).toHaveCount(1)
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })

  test('documento comum abre editável', async () => {
    const target = join(folder, 'simples.docx')
    await writeFile(target, await docxWithoutExtras())
    await stubDialogs(session.app, { open: target, messageBox: 1 })

    await menu(session, 'open')

    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toBeHidden()
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })
})
