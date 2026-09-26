/// <reference lib="dom" />
import { execFile } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithNamedStyles, docxWithReferences, entryOf } from './fixtures.js'

/**
 * Referências (M8): painel de navegação, marcadores, sumário, legendas e
 * referências cruzadas.
 *
 * A preferência é mudada por `window.api.preferences.set`, e não pelo menu
 * nativo, pelo mesmo motivo de `exibir.spec.ts`: o Playwright não alcança o menu
 * do sistema, e o caminho depois do clique é o mesmo.
 */
async function setPreference(session: Session, patch: Record<string, unknown>): Promise<void> {
  await session.window.evaluate(async (value) => {
    const api = (window as unknown as { api: { preferences: { set: (p: unknown) => Promise<unknown> } } }).api
    await api.preferences.set(value)
  }, patch)
}

test.describe('painel de navegação', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-referencias-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('lista os títulos ao vivo, leva a eles e destaca o da seção do cursor', async () => {
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()

    await session.window.keyboard.press('Control+Alt+1')
    await session.window.keyboard.type('Capítulo um')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Texto do capítulo.')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Control+Alt+2')
    await session.window.keyboard.type('Seção dois')

    await setPreference(session, { navigationPane: true })
    const pane = session.window.getByRole('navigation', { name: 'Navegação' })
    await expect(pane).toBeVisible()

    const entries = pane.getByRole('button', { name: /^Nível/ })
    await expect(entries).toHaveText(['Capítulo um', 'Seção dois'])
    // O cursor está no fim do segundo título: é a seção dele que se destaca.
    await expect(entries.nth(1)).toHaveAttribute('aria-current', 'location')

    await entries.nth(0).click()
    await expect(entries.nth(0)).toHaveAttribute('aria-current', 'location')
    // O cursor foi para o título: digitar escreve nele.
    await session.window.keyboard.type('>')
    await expect(session.window.locator('.ProseMirror h1')).toHaveText('>Capítulo um')

    await setPreference(session, { navigationPane: false })
    await expect(pane).toHaveCount(0)
  })

  test('no documento travado também leva ao título', async () => {
    // O comentário trava a edição. O título é o `Ttulo1` do Word em português, que
    // o painel reconhece pelo nome interno `heading 1`.
    const origem = join(folder, 'travado.docx')
    const p = (style: string, text: string): string =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    await writeFile(
      origem,
      await docxWithNamedStyles(
        Array.from({ length: 40 }, (_, index) => p('Normal', `Parágrafo ${index + 1}.`)).join('') +
          p('Ttulo1', 'Conclusão') +
          '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Comentado.</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>',
      ),
    )
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.banner--readonly')).toBeVisible()

    await setPreference(session, { navigationPane: true })
    const entries = session.window
      .getByRole('navigation', { name: 'Navegação' })
      .getByRole('button', { name: /^Nível/ })
    await expect(entries).toHaveText(['Relatório anual', 'Conclusão'])

    await entries.nth(1).click()
    await expect(entries.nth(1)).toHaveAttribute('aria-current', 'location')
    // Rolou até o título, que estava folhas abaixo.
    await expect(session.window.locator('.ProseMirror h1', { hasText: 'Conclusão' })).toBeInViewport()
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
  })
})

/**
 * Marcadores sem sumário nem campos: os dois ainda travavam o documento quando
 * este teste nasceu, e aqui o que se quer é editar.
 */
const BOOKMARKS_BODY =
  '<w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc100"/><w:r><w:t>Introdução</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
  '<w:p><w:bookmarkStart w:id="1" w:name="Resumo"/><w:r><w:t xml:space="preserve">O resumo começa aqui </w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>e termina aqui.</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>' +
  '<w:p><w:r><w:t xml:space="preserve">Veja o </w:t></w:r><w:hyperlink w:anchor="Resumo" w:history="1"><w:r><w:t>resumo</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> e a conclusão.</w:t></w:r></w:p>'

test.describe('marcadores', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-marcadores-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('lista, vai para, adiciona e grava sem perder os ocultos', async () => {
    const origem = join(folder, 'marcadores.docx')
    const destino = join(folder, 'saida.docx')
    await writeFile(origem, await docxWithReferences(BOOKMARKS_BODY))
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('O resumo começa aqui')
    await expect(editor).toHaveAttribute('contenteditable', 'true')

    await menu(session, 'insert-bookmark')
    const dialog = session.window.getByRole('dialog', { name: 'Marcadores' })
    const list = dialog.getByRole('listbox', { name: 'Marcadores do documento' })
    // O oculto do sumário só aparece com a caixa marcada.
    await expect(list.getByRole('option')).toHaveText(['Resumo'])
    await dialog.getByLabel('Marcadores ocultos').check()
    await expect(list.getByRole('option')).toHaveText(['_Toc100', 'Resumo'])

    await list.getByRole('option', { name: 'Resumo' }).click()
    await dialog.getByRole('button', { name: 'Ir para' }).click()
    await expect(dialog).toHaveCount(0)
    await session.window.keyboard.type('>')
    await expect(editor.locator('p', { hasText: 'começa aqui' })).toHaveText('>O resumo começa aqui ')

    // Um marcador novo sobre o ponto final do último parágrafo.
    const last = editor.locator('p', { hasText: 'e a conclusão' })
    await last.dblclick({ position: { x: 5, y: 5 } })
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+ArrowLeft')
    await menu(session, 'insert-bookmark')
    await dialog.getByLabel('Nome do marcador').fill('com espaço')
    await expect(dialog.getByRole('button', { name: 'Adicionar' })).toBeDisabled()
    await dialog.getByLabel('Nome do marcador').fill('Conclusao')
    await dialog.getByRole('button', { name: 'Adicionar' }).click()
    await expect(dialog).toHaveCount(0)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toMatch(/w:name="Conclusao"/)
    // O oculto do título e a ponta que mora no corpo, entre dois parágrafos.
    expect(corpo).toMatch(/w:name="_Toc100"/)
    expect(corpo).toMatch(/<\/w:p><w:bookmarkEnd w:id="1" ?\/>/)
    expect(corpo).toMatch(/w:anchor="Resumo"/)
  })

  test('Ctrl+clique no link interno leva ao marcador', async () => {
    const origem = join(folder, 'link.docx')
    await writeFile(origem, await docxWithReferences(BOOKMARKS_BODY))
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')

    await editor.locator('a', { hasText: 'resumo' }).click({ modifiers: ['Control'] })
    await session.window.keyboard.type('>')
    await expect(editor.locator('p', { hasText: 'começa aqui' })).toHaveText('>O resumo começa aqui ')
  })

  test('link para um título cria o marcador oculto em volta dele', async () => {
    const origem = join(folder, 'titulo.docx')
    const destino = join(folder, 'titulo-saida.docx')
    await writeFile(
      origem,
      await docxWithReferences(BOOKMARKS_BODY.replace(/<w:bookmark(Start|End)[^>]*\/>/g, '')),
    )
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')

    // Cursor no fim do último parágrafo, sem seleção: o link leva o texto do título.
    await editor.locator('p', { hasText: 'e a conclusão' }).click()
    await session.window.keyboard.press('End')
    await session.window.getByRole('button', { name: 'Link' }).click()
    const dialog = session.window.getByRole('dialog', { name: 'Inserir link' })
    await dialog.getByLabel('Lugar neste documento').selectOption({ label: 'Introdução' })
    await dialog.getByRole('button', { name: 'Aplicar' }).click()
    await expect(editor.locator('a', { hasText: 'Introdução' })).toHaveAttribute('href', /^#_Ref\d+$/)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    const nome = /w:anchor="(_Ref\d+)"/.exec(corpo)?.[1]
    expect(nome).toBeDefined()
    expect(corpo).toContain(`w:name="${nome ?? ''}"`)
  })
})

/** O texto de uma página do PDF, pelo `pdftotext` do sistema — nulo sem ele. */
async function textoDaPagina(caminho: string, pagina: number): Promise<string | null> {
  try {
    const { stdout } = await promisify(execFile)('pdftotext', [
      '-f',
      String(pagina),
      '-l',
      String(pagina),
      '-layout',
      caminho,
      '-',
    ])
    return stdout
  } catch {
    return null
  }
}

async function pdfPronto(caminho: string): Promise<boolean> {
  try {
    return (await readFile(caminho)).toString('latin1').includes('%%EOF')
  } catch {
    return false
  }
}

test.describe('sumário', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-sumario-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('insere com o número da página de cada título, atualiza e sai igual no PDF e no .docx', async () => {
    const pdf = join(folder, 'sumario.pdf')
    const docx = join(folder, 'sumario.docx')
    await stubDialogs(session.app, { save: pdf, messageBox: 1 })
    await menu(session, 'new-document')
    const editor = session.window.locator('.ProseMirror')
    await editor.click()

    await session.window.keyboard.type('Capa.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.press('Control+Alt+1')
    await session.window.keyboard.type('Introdução')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Texto.')
    await menu(session, 'insert-page-break')
    await session.window.keyboard.press('Control+Alt+2')
    await session.window.keyboard.type('Escopo')
    await expect(session.window.locator('.paper')).toHaveCount(3)

    // O sumário entra antes do bloco do cursor, na primeira folha.
    await editor.locator('p', { hasText: 'Capa.' }).click()
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-table-of-contents')
    const toc = editor.locator('.toc')
    await expect(toc).toContainText('Sumário')
    // O número é o da folha em que o título caiu, depois de o sumário ocupar a
    // primeira — o segundo passe da paginação.
    await expect(toc.locator('.field')).toHaveText(['2', '3'])
    await expect(toc.locator('a').first()).toHaveAttribute('href', /^#_Toc\d+$/)

    // Um título novo, e "Atualizar sumário" o acrescenta.
    // Pelo painel de navegação: leva o cursor ao título sem depender de onde o
    // clique cai na folha.
    await setPreference(session, { navigationPane: true })
    await session.window
      .getByRole('navigation', { name: 'Navegação' })
      .getByRole('button', { name: /Escopo/ })
      .click()
    await session.window.keyboard.press('End')
    // O `End` é do navegador, e o ProseMirror só o lê no `selectionchange`; uma
    // pessoa nunca aperta a tecla seguinte no mesmo milissegundo, o teste sim.
    await session.window.waitForTimeout(100)
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Control+Alt+1')
    await session.window.keyboard.type('Conclusão')
    await menu(session, 'update-table-of-contents')

    await expect(toc.locator('p', { hasText: 'Conclusão' })).toHaveCount(1)
    await expect(toc.locator('.field')).toHaveText(['2', '3', '3'])

    // O papel: a mesma primeira folha, com as entradas e os números.
    await menu(session, 'export-pdf')
    await expect.poll(() => pdfPronto(pdf), { timeout: 30_000 }).toBe(true)
    const primeira = await textoDaPagina(pdf, 1)
    if (primeira !== null) {
      expect(primeira).toMatch(/Introdução[ .]*2/)
      expect(primeira).toMatch(/Escopo[ .]*3/)
      expect(primeira).toMatch(/Conclusão[ .]*3/)
    }

    await stubDialogs(session.app, { save: docx, messageBox: 1 })
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(docx, 'word/document.xml')
    expect(corpo).toContain('w:val="Table of Contents"')
    expect(corpo).toMatch(/TOC \\o "1-3" \\h/)
    expect(corpo.match(/PAGEREF _Toc\d+/g)).toHaveLength(3)
    expect(
      corpo.match(/w:bookmarkStart[^>]*w:name="_Toc\d+"|w:name="_Toc\d+"[^>]*w:bookmarkStart/g)?.length ?? 0,
    ).toBeGreaterThan(0)
    const estilos = await entryOf(docx, 'word/styles.xml')
    expect(estilos).toContain('w:val="toc 1"')
    expect(estilos).toContain('w:val="TOC Heading"')
  })

  test('o sumário do Word abre editável, mostra as entradas e volta intacto', async () => {
    const origem = join(folder, 'word.docx')
    const destino = join(folder, 'word-saida.docx')
    await writeFile(origem, await docxWithReferences())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')

    // Campos e sumário representados não travam mais o documento.
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toHaveCount(0)
    const toc = editor.locator('.toc')
    await expect(toc).toContainText('Sumário')
    await expect(toc.locator('.field')).toHaveText(['1', '1'])

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const antes = await entryOf(origem, 'word/document.xml')
    const depois = await entryOf(destino, 'word/document.xml')
    const sumario = (xml: string): string => /<w:sdt>.*<\/w:sdt>/.exec(xml)?.[0] ?? ''
    expect(sumario(depois)).not.toBe('')
    // Preservado sem ninguém pedir "Atualizar": o mesmo XML, peça por peça.
    // A ordem dos atributos é a do SDK, que reserializa o corpo; o conteúdo não muda.
    const normal = (xml: string): string =>
      sumario(xml)
        .replace(/ \/>/g, '/>')
        .replace(/ w:history="1"/g, '')
    expect(normal(depois)).toBe(normal(antes))
  })
})

test.describe('legendas e referências cruzadas', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-legendas-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('legenda nova, F9 renumera, e a referência cruzada cita o número e o título', async () => {
    const origem = join(folder, 'legendas.docx')
    const destino = join(folder, 'legendas-saida.docx')
    await writeFile(origem, await docxWithReferences())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toContainText('Como mostra a Figura 1')

    // Uma legenda nova logo abaixo de "Escopo", antes da Figura 1 do arquivo.
    await setPreference(session, { navigationPane: true })
    await session.window
      .getByRole('navigation', { name: 'Navegação' })
      .getByRole('button', { name: /Escopo/ })
      .click()
    await menu(session, 'insert-caption')
    const legenda = session.window.getByRole('dialog', { name: 'Legenda' })
    await legenda.getByLabel('Rótulo').selectOption('Figura')
    await legenda.getByLabel('Texto depois do número').fill('— Fluxo')
    await legenda.getByRole('button', { name: 'Inserir' }).click()
    await expect(editor.locator('p', { hasText: 'Fluxo' })).toHaveText('Figura 1 — Fluxo')

    // F9 com o cursor parado: o documento inteiro. A legenda do arquivo vira 2,
    // e a referência a ela acompanha.
    await menu(session, 'update-fields')
    await expect(editor.locator('p', { hasText: 'Arquitetura' })).toHaveText('Figura 2 — Arquitetura')
    await expect(editor).toContainText('Como mostra a Figura 2, na página 1')

    // Referências cruzadas no fim do último parágrafo: o número da figura do
    // arquivo e o texto de um título.
    await editor.locator('p', { hasText: 'Veja o' }).click()
    await session.window.keyboard.press('End')
    await session.window.waitForTimeout(100)
    await menu(session, 'insert-cross-reference')
    const ref = session.window.getByRole('dialog', { name: 'Referência cruzada' })
    await ref.getByLabel('Tipo').selectOption('caption:Figura')
    await ref.getByLabel('Para qual').selectOption({ label: 'Figura 2 — Arquitetura' })
    await ref.getByLabel('Inserir referência a').selectOption('number')
    await ref.getByRole('button', { name: 'Inserir' }).click()

    await menu(session, 'insert-cross-reference')
    await ref.getByLabel('Tipo').selectOption('heading')
    await ref.getByLabel('Para qual').selectOption({ label: 'Introdução' })
    await ref.getByRole('button', { name: 'Inserir' }).click()
    await expect(editor.locator('p', { hasText: 'Veja o' })).toHaveText(
      'Como mostra a Figura 2, na página 1. Veja o resumo.2Introdução',
    )

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo.match(/SEQ Figura/g)).toHaveLength(2)
    expect(corpo).toMatch(/REF _Ref\d+ \\h/)
    expect(corpo).toMatch(/REF _Ref200 \\h/)
    // O número novo da figura do arquivo foi para o resultado do campo.
    expect(corpo).toMatch(/<w:t[^>]*>Figura 2<\/w:t>/)
  })
})
