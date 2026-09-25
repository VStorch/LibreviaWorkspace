/// <reference lib="dom" />
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, entryOf } from './fixtures.js'

/**
 * O painel de estilos: ver, aplicar, modificar, criar e limpar.
 *
 * Responde na tela duas perguntas que o programa não sabia responder — quais
 * estilos o documento tem, e qual é o do parágrafo onde está o cursor — e deixa
 * agir sobre a resposta. O que se prova aqui é o que só o aplicativo inteiro
 * mostra: a tela mudando com o estilo, o desfazer, e o estilo modificado chegando
 * ao `.docx` e voltando dele.
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
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Título 1' })
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

  test('aplicar pelo painel troca o bloco, e desfazer o devolve', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Relatório')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.locator('.styles-list__item', { hasText: 'Título 1' }).click()
    await panel.getByRole('button', { name: 'Aplicar' }).click()

    await expect(editor.locator('h1')).toHaveText('Relatório')
    await expect(panel).toContainText('Parágrafo do cursor: Título 1')

    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('h1')).toHaveCount(0)
    await expect(editor.locator('p').first()).toHaveText('Relatório')
  })

  test('modificar o estilo muda a tela na hora', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Corpo do texto.')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel
      .locator('.styles-list__item')
      .filter({ has: session.window.locator('.styles-list__name', { hasText: /^Normal$/ }) })
      .click()
    await panel.getByRole('button', { name: 'Modificar…' }).click()

    // O embutido não muda de nome: é por ele que o Word o reconhece.
    await expect(panel.getByRole('textbox', { name: 'Nome' })).toBeDisabled()
    await panel.getByRole('spinbutton', { name: 'Tamanho (pt)' }).fill('20')
    await panel.getByRole('button', { name: 'OK' }).click()

    // 20 pt em pixels de CSS: a regra do estilo foi regerada, e o parágrafo sem
    // formatação direta a segue.
    const size = await editor
      .locator('p')
      .first()
      .evaluate((element) => getComputedStyle(element).fontSize)
    expect(Number.parseFloat(size)).toBeCloseTo(20 * (96 / 72), 0)
    await expect(session.window.locator('.statusbar__state')).not.toHaveText('Salvo')
  })

  test('Enter no fim do título abre o estilo seguinte', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Capítulo')
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Título 1' })

    await editor.locator('h1').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Texto depois do título.')

    await expect(editor.locator('p').first()).toHaveText('Texto depois do título.')
    await expect(session.window.getByRole('combobox', { name: 'Estilo' })).toHaveValue('Normal')
  })

  test('limpar a formatação tira o direto e deixa o estilo', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type('Negrito à mão')
    await expect(editor.locator('strong')).toHaveCount(1)

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.getByRole('button', { name: 'Limpar formatação' }).click()

    await expect(editor.locator('strong')).toHaveCount(0)
    await expect(editor.locator('p').first()).toHaveText('Negrito à mão')
  })

  test('o estilo modificado vai para o .docx e volta dele', async () => {
    const origem = join(folder, 'estilos.docx')
    await writeFile(origem, await docxWithNamedStyles())

    await stubDialogs(session.app, { open: origem, save: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Um trecho citado.')

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.locator('.styles-list__item', { hasText: 'Citação recuada' }).click()
    await panel.getByRole('button', { name: 'Modificar…' }).click()
    await panel.getByRole('spinbutton', { name: 'Tamanho (pt)' }).fill('18')
    await panel.getByRole('button', { name: 'OK' }).click()
    await panel.getByRole('button', { name: 'Fechar' }).click()

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // Só o estilo mudou no arquivo: o parágrafo continua só apontando para ele.
    expect(await entryOf(origem, 'word/styles.xml')).toContain('<w:sz w:val="36"')
    expect(await entryOf(origem, 'word/document.xml')).not.toContain('w:sz')

    await menu(session, 'open')
    const quote = editor.locator('[data-style-id="Citao"]')
    await expect(quote).toHaveText('Um trecho citado.')
    const size = await quote.evaluate((element) => getComputedStyle(element).fontSize)
    expect(Number.parseFloat(size)).toBeCloseTo(18 * (96 / 72), 0)
  })

  test('o texto importado segue o estilo: modificado, e de título a Normal', async () => {
    // Mede o **trecho** — o elemento que envolve o texto —, e não o parágrafo:
    // com a fonte do estilo presa numa marca do trecho, o parágrafo mudaria e o
    // texto não.
    const origem = join(folder, 'estilos.docx')
    await writeFile(origem, await docxWithNamedStyles())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Um trecho citado.')

    const look = (text: string) =>
      editor.evaluate((root, needle) => {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
        for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
          if (node.textContent?.includes(needle) !== true) continue
          const style = getComputedStyle(node.parentElement!)
          return { size: Number.parseFloat(style.fontSize), weight: style.fontWeight }
        }
        return null
      }, text)

    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()
    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await panel.locator('.styles-list__item', { hasText: 'Citação recuada' }).click()
    await panel.getByRole('button', { name: 'Modificar…' }).click()
    await panel.getByRole('spinbutton', { name: 'Tamanho (pt)' }).fill('18')
    await panel.getByRole('button', { name: 'OK' }).click()
    await panel.getByRole('button', { name: 'Fechar' }).click()

    expect((await look('Um trecho citado.'))?.size).toBeCloseTo(18 * (96 / 72), 0)

    // O título do arquivo é negrito de 16 pt pelo estilo; trocado por Normal, o
    // texto volta a 11 pt e sem negrito — nada do título ficou preso no trecho.
    expect((await look('Relatório anual'))?.weight).toBe('700')
    await editor.getByText('Relatório anual').click()
    await session.window.getByRole('combobox', { name: 'Estilo' }).selectOption({ label: 'Normal' })
    const normal = await look('Relatório anual')
    expect(normal?.size).toBeCloseTo(11 * (96 / 72), 0)
    expect(normal?.weight).toBe('400')
  })

  test('o foco circula dentro do painel, e Escape continua fechando', async () => {
    await menu(session, 'new-document')
    await session.window.getByRole('button', { name: 'Estilos do documento' }).click()

    const panel = session.window.getByRole('dialog', { name: 'Estilos' })
    await expect(panel.getByRole('button', { name: 'Fechar' })).toBeFocused()

    // Sem a circulação, o Tab daqui caía no texto do documento — e dali o
    // Escape não fechava mais o painel, porque quem escuta a tecla é ele.
    await panel.press('Tab')
    await expect(panel.getByRole('combobox', { name: 'Mostrar' })).toBeFocused()

    await panel.press('Shift+Tab')
    await expect(panel.getByRole('button', { name: 'Fechar' })).toBeFocused()

    await panel.press('Escape')
    await expect(panel).toBeHidden()
  })
})
