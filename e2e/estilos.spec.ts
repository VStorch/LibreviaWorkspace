/// <reference lib="dom" />
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles } from './fixtures.js'

/**
 * O painel de estilos, só de leitura.
 *
 * Responde na tela duas perguntas que o programa não sabia responder: quais
 * estilos o documento tem, e qual é o do parágrafo onde está o cursor. Num
 * `.docx` corporativo quase toda a formatação mora em estilos, e quem abria um
 * aqui via o resultado sem nunca ver a regra.
 *
 * Aplicar, criar e modificar são das entregas seguintes — e é isso que o último
 * teste protege: o painel não pode ganhar um botão que mexa no documento sem que
 * alguém decida que ele pode.
 */
test.describe('painel de estilos', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-estilos-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('o documento novo mostra os estilos embutidos e o do cursor', async () => {
    await menu(session, 'new-document')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })

    // O parágrafo de um documento novo não aponta estilo nenhum: quem vale é o
    // estilo padrão do documento, e é ele que o painel precisa nomear.
    await expect(panel).toContainText('Parágrafo do cursor: Normal')
    await expect(panel).toContainText('Título 1')

    // Virar título muda a resposta **ao vivo**, sem fechar o painel: o cursor
    // continua onde estava, e o seletor da barra devolve o foco ao documento.
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption('1')
    await expect(panel).toContainText('Parágrafo do cursor: Título 1')
  })

  test('mostra os estilos que o .docx traz, com o nome que o autor deu', async () => {
    const origem = join(folder, 'estilos.docx')
    await writeFile(origem, await docxWithNamedStyles())

    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Um trecho citado.')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })

    // O estilo criado por quem escreveu o documento aparece pelo nome dele, e o
    // título pelo nome interno traduzido — o id (`Ttulo1`) é do arquivo, não da
    // tela.
    await expect(panel).toContainText('Citação recuada')
    await expect(panel).toContainText('Título 1')

    // A maquinaria do Word fica fora da lista: o `w:semiHidden` existe para isso.
    await expect(panel).not.toContainText('Default Paragraph Font')

    // O cursor abre no primeiro bloco, que é o título — e o estilo dele vem do
    // arquivo, com id em português.
    await expect(panel).toContainText('Parágrafo do cursor: Título 1')

    // Só os de caractere: o filtro é o que torna legível a lista de um documento
    // do Word, que declara dezenas.
    await panel.getByRole('combobox', { name: 'Mostrar' }).selectOption('character')
    await expect(panel).not.toContainText('Citação recuada')
  })

  test('não oferece nada que altere o documento, e fecha com Escape', async () => {
    await menu(session, 'new-document')
    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()

    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    // Um botão só: fechar. Aplicar um estilo reescreve o `w:pStyle` de um
    // parágrafo e modificá-lo reescreve `word/styles.xml` — as duas coisas mexem
    // no arquivo de quem confia neste programa, e cada uma tem a sua entrega.
    await expect(panel.getByRole('button')).toHaveCount(1)
    await expect(panel.getByRole('button', { name: 'Fechar' })).toBeFocused()

    await panel.press('Escape')
    await expect(panel).toBeHidden()
  })
})
