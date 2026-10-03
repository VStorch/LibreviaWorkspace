import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithCommentThread, entryOf } from './fixtures.js'

/** Clica num item do menu nativo, pelo rótulo do menu e do item. */
async function clickMenuItem(session: Session, menuLabel: string, itemLabel: string): Promise<void> {
  await session.app.evaluate(
    ({ Menu }, [top, item]) => {
      const submenu = Menu.getApplicationMenu()?.items.find((entry) => entry.label === top)?.submenu
      submenu?.items.find((entry) => entry.label === item)?.click()
    },
    [menuLabel, itemLabel] as const,
  )
}

async function temPdftotext(): Promise<boolean> {
  try {
    await promisify(execFile)('pdftotext', ['-v'])
    return true
  } catch {
    return false
  }
}

async function textoDoPdf(caminho: string): Promise<string> {
  try {
    const { stdout } = await promisify(execFile)('pdftotext', ['-layout', caminho, '-'])
    return stdout
  } catch {
    return ''
  }
}

/**
 * Comentários: lidos, mostrados ao lado da folha e devolvidos ao arquivo quando
 * o parágrafo que os ancora é editado; criados, respondidos, resolvidos e
 * excluídos no painel, e gravados de volta; o cursor em volta das âncoras,
 * recortar e colar, navegar, esconder o painel, o autor e o papel.
 */
test.describe('comentários', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-comentarios-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('a conversa aparece ao lado do trecho, e editar o parágrafo não a perde', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(session.window.locator('.banner--readonly')).toBeHidden()

    // Uma conversa por cartão: a resposta mora dentro do da pergunta.
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await expect(conversa).toContainText('Ana')
    await expect(conversa.locator('.comment-card__replies')).toContainText('Conferido na planilha.')

    // A resolvida vem recolhida: o autor e o aviso, sem o texto.
    const resolvida = cartoes.filter({ hasText: 'Resolvido' })
    await expect(resolvida).toContainText('Carla')
    await expect(resolvida).not.toContainText('Trocar o título.')

    // Na altura do trecho que comenta.
    const trecho = editor.locator('.comment-range').first()
    await expect(trecho).toHaveText('doze mil reais')
    await expect
      .poll(async () => {
        const [caixaTrecho, caixaCartao] = await Promise.all([trecho.boundingBox(), conversa.boundingBox()])
        return Math.abs(caixaTrecho!.y - caixaCartao!.y)
      })
      .toBeLessThan(30)

    // Clicar no cartão seleciona o trecho e o realça.
    await conversa.click()
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')

    // Editar o parágrafo comentado e gravar.
    await editor.getByText('por ano.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Revisado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await expect(session.window.locator('.banner--notice')).toBeHidden()

    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('Revisado.')
    for (const id of ['0', '1', '2']) {
      expect(corpo).toContain(`<w:commentRangeStart w:id="${id}"`)
      expect(corpo).toContain(`<w:commentReference w:id="${id}"`)
    }

    // Reaberto, o cartão continua lá, com a resposta.
    await menu(session, 'open')
    await expect(editor).toContainText('Revisado.')
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Conferir o valor.' })).toContainText('Conferido na planilha.')
  })
  test('criar, responder, resolver e excluir, e o arquivo e o rascunho guardam tudo', async () => {
    const origem = join(pasta, 'ata.docx')
    const rascunho = join(pasta, 'ata.sdoc')
    const destino = join(pasta, 'volta.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, save: rascunho, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)

    // Novo comentário sobre "Fim.": o cartão abre com a caixa de texto em foco.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    const caixa = session.window.locator('.comment-composer__field')
    await expect(caixa).toBeFocused()
    await session.window.keyboard.type('Revisar o fim.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes).toHaveCount(3)
    await expect(cartoes.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
    await expect(editor.locator('.comment-range').filter({ hasText: 'Fim.' })).toHaveCount(1)

    // Esc desiste do cartão novo, e as pontas saem com ele.
    await editor.getByText('Ata da reunião').click()
    await menu(session, 'insert-comment')
    await expect(caixa).toBeFocused()
    await session.window.keyboard.press('Escape')
    await expect(cartoes).toHaveCount(3)

    // Responder e resolver a conversa da Ana.
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await conversa.click()
    await conversa.getByRole('button', { name: 'Responder' }).click()
    await session.window.keyboard.type('Fechado.')
    await session.window.keyboard.press('Control+Enter')
    await expect(conversa.locator('.comment-card__replies')).toContainText('Fechado.')
    await conversa.getByRole('button', { name: 'Resolver' }).click()
    await expect(conversa).toHaveClass(/comment-card--done/)

    // Excluir a conversa resolvida da Carla.
    const resolvida = cartoes.filter({ hasText: 'Carla' })
    await resolvida.click()
    await resolvida.getByRole('button', { name: 'Excluir' }).click()
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Carla' })).toHaveCount(0)

    // Desfazer a exclusão devolve o cartão; refazer o tira de novo.
    await editor.click()
    await session.window.keyboard.press('Control+Z')
    await expect(cartoes).toHaveCount(3)
    await session.window.keyboard.press('Control+Shift+Z')
    await expect(cartoes).toHaveCount(2)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const comentarios = await entryOf(origem, 'word/comments.xml')
    expect(comentarios).toContain('Revisar o fim.')
    expect(comentarios).toContain('Fechado.')
    expect(comentarios).not.toContain('Trocar o título.')
    expect(await entryOf(origem, 'word/commentsExtended.xml')).toMatch(/w15:paraId="10000000" w15:done="1"/)
    expect(await entryOf(origem, 'word/document.xml')).not.toContain('w:id="2"')

    // O rascunho guarda o mesmo estado, e o .docx gravado dele também.
    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await stubDialogs(session.app, { open: rascunho, save: destino })
    await menu(session, 'close-file')
    await expect(session.window.locator('.home')).toBeVisible()
    await menu(session, 'open')
    await expect(cartoes).toHaveCount(2)
    await expect(cartoes.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
    await expect(cartoes.filter({ hasText: 'Ana' })).toHaveClass(/comment-card--done/)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await stubDialogs(session.app, { open: destino })
    await menu(session, 'open')
    await expect(cartoes).toHaveCount(2)
    const ana = cartoes.filter({ hasText: 'Ana' })
    await expect(ana).toHaveClass(/comment-card--done/)
    await ana.click()
    await expect(ana.locator('.comment-card__replies')).toContainText('Conferido na planilha.')
    await expect(ana.locator('.comment-card__replies')).toContainText('Fechado.')
    await expect(cartoes.filter({ hasText: 'Revisar o fim.' })).toBeVisible()
  })
  test('o comentário de ponto no fim do parágrafo não desencontra a seleção', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)

    // Comentário de ponto no fim de "Fim.": só o `commentEnd`, encostado no fim.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await menu(session, 'insert-comment')
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.type('Ponto.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes).toHaveCount(3)

    // Só o último caractere selecionado não é o parágrafo: apagá-lo deixa o comentário.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+ArrowLeft')
    await session.window.keyboard.press('Delete')
    await expect(editor.locator('p').last()).toHaveText('Fim')
    await expect(cartoes).toHaveCount(3)
    await session.window.keyboard.type('.')

    // `End` e `Shift+Home` depressa selecionam a linha, e um `Backspace` só a apaga.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type('Novo')
    await expect(editor.locator('p').last()).toHaveText('Novo')

    // Sem seleção, o `Backspace` passa por cima da âncora — o fim do comentário
    // da Carla — e apaga o caractere, e a conversa fica.
    await editor.getByText('Título provisório').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Backspace')
    await expect(editor.getByText('Título provisóri', { exact: true })).toBeVisible()
    await expect(cartoes.filter({ hasText: 'Carla' })).toHaveCount(1)
  })
  test('recortar e colar leva o comentário; copiar e colar não o repete', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    const carla = cartoes.filter({ hasText: 'Carla' })
    await expect(carla).toHaveCount(1)

    // A linha inteira leva as pontas encostadas nela: recortada, a conversa sai.
    await editor.getByText('Título provisório').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Shift+End')
    await session.window.keyboard.press('Control+X')
    await expect(carla).toHaveCount(0)

    // Colada no fim, volta com o cartão.
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.press('Control+V')
    await expect(carla).toHaveCount(1)
    await expect(editor.locator('[data-comment-start][data-cid="2"]')).toHaveCount(1)

    // Copiada e colada de novo, o texto se repete e a conversa não.
    await editor.getByText('Título provisório').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.press('Shift+End')
    await session.window.keyboard.press('Control+C')
    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Control+V')
    await expect(editor.getByText('Título provisório')).toHaveCount(2)
    await expect(editor.locator('[data-comment-start][data-cid="2"]')).toHaveCount(1)
    await expect(carla).toHaveCount(1)

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    expect(await entryOf(origem, 'word/comments.xml')).toContain('Trocar o título.')
    expect((await entryOf(origem, 'word/document.xml')).match(/<w:commentRangeStart w:id="2"/g)).toHaveLength(
      1,
    )
  })

  test('próximo e anterior escolhem a conversa pela ordem do texto', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const ativo = session.window.locator('.comment-card--active')
    await expect(session.window.locator('.comment-card')).toHaveCount(2)

    // A resposta do Bruno tem pontas próprias no arquivo, mas não é conversa.
    await menu(session, 'next-comment')
    await expect(ativo).toContainText('Ana')
    await expect(editor.locator('.comment-range--active')).toHaveText('doze mil reais')
    await menu(session, 'next-comment')
    await expect(ativo).toContainText('Carla')
    await menu(session, 'next-comment')
    await expect(ativo).toContainText('Ana')
    await menu(session, 'previous-comment')
    await expect(ativo).toContainText('Carla')

    // O cursor foi junto: o que se digita cai no trecho escolhido.
    await expect(editor).toBeFocused()
  })

  test('esconder o painel tira os cartões e o realce, e o arquivo guarda tudo', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const painel = session.window.locator('.comments-pane')
    await expect(painel).toHaveCount(1)
    await expect(editor.locator('.comment-range')).not.toHaveCount(0)

    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(painel).toHaveCount(0)
    await expect(editor.locator('.comment-range')).toHaveCount(0)

    await editor.getByText('por ano.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Revisado.')
    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('<w:commentRangeStart w:id="0"')
    expect(corpo).toContain('<w:commentRangeStart w:id="2"')
    expect(await entryOf(origem, 'word/comments.xml')).toContain('Conferir o valor.')

    // A preferência fica: reaberto o documento, o painel continua escondido.
    await menu(session, 'open')
    await expect(editor).toContainText('Revisado.')
    await expect(painel).toHaveCount(0)

    // Inserir um comentário traz o painel de volta.
    await editor.getByText('Fim.').click()
    await menu(session, 'insert-comment')
    await expect(painel).toHaveCount(1)
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.press('Escape')

    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(painel).toHaveCount(0)
    await clickMenuItem(session, 'Exibir', 'Comentários')
    await expect(painel).toHaveCount(1)
    await expect(editor.locator('.comment-range').first()).toBeVisible()
  })

  test('o nome do autor assina o comentário novo, e a resposta se edita', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.ProseMirror')
    const cartoes = session.window.locator('.comment-card')
    await expect(cartoes).toHaveCount(2)

    await menu(session, 'author-name')
    const dialogo = session.window.getByRole('dialog', { name: 'Nome do autor' })
    const nome = dialogo.getByRole('textbox', { name: 'Nome' })
    await expect(nome).toBeFocused()
    await nome.fill('Zé da Silva')
    await dialogo.getByRole('button', { name: 'Salvar' }).click()
    await expect(dialogo).toHaveCount(0)

    await editor.getByText('Fim.').click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('Shift+Home')
    await menu(session, 'insert-comment')
    await expect(session.window.locator('.comment-composer__field')).toBeFocused()
    await session.window.keyboard.type('Assinado.')
    await session.window.getByRole('button', { name: 'Comentar' }).click()
    await expect(cartoes.filter({ hasText: 'Assinado.' })).toContainText('Zé da Silva')

    // A resposta do Bruno, editada no cartão da Ana.
    const conversa = cartoes.filter({ hasText: 'Conferir o valor.' })
    await conversa.click()
    const respostas = conversa.locator('.comment-card__replies')
    await respostas.getByRole('button', { name: 'Editar' }).click()
    const caixa = respostas.locator('.comment-composer__field')
    await expect(caixa).toBeFocused()
    await expect(caixa).toHaveValue('Conferido na planilha.')
    await caixa.fill('Conferido duas vezes.')
    await session.window.keyboard.press('Control+Enter')
    await expect(respostas).toContainText('Conferido duas vezes.')
    await expect(respostas).not.toContainText('Conferido na planilha.')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const comentarios = await entryOf(origem, 'word/comments.xml')
    expect(comentarios).toMatch(
      /w:author="Zé da Silva"[^>]*w:initials="ZDS"|w:initials="ZDS"[^>]*w:author="Zé da Silva"/,
    )
    expect(comentarios).toContain('Conferido duas vezes.')
    expect(comentarios).not.toContain('Conferido na planilha.')
  })

  test('o PDF sai sem os comentários', async () => {
    test.skip(!(await temPdftotext()), 'pdftotext não instalado')
    const origem = join(pasta, 'ata.docx')
    const destino = join(pasta, 'ata.pdf')
    await writeFile(origem, await docxWithCommentThread())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.comment-card')).toHaveCount(2)

    await menu(session, 'export-pdf')
    await expect.poll(() => textoDoPdf(destino), { timeout: 30_000 }).toContain('doze mil reais')
    const texto = await textoDoPdf(destino)
    for (const fora of ['Conferir o valor.', 'Conferido na planilha.', 'Ana', 'Bruno', 'Carla', 'Resolvido'])
      expect(texto).not.toContain(fora)
  })
})
