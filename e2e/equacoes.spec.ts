import { execFile } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { EQUATION_AFTER, EQUATION_BEFORE, docxWithEquations, entryOf } from './fixtures.js'

async function textoDoPdf(caminho: string): Promise<string> {
  try {
    const { stdout } = await promisify(execFile)('pdftotext', [caminho, '-'])
    return stdout
  } catch {
    return ''
  }
}

async function temPdftotext(): Promise<boolean> {
  try {
    await promisify(execFile)('pdftotext', ['-v'])
    return true
  } catch {
    return false
  }
}

/** Os `m:oMath`/`m:oMathPara` do XML, na ordem — o que tem de voltar como veio. */
function equacoesDo(xml: string): string[] {
  // O `word/document.xml` inteiro é serializado de novo pelo SDK, que fecha o
  // elemento vazio com " />": é a única diferença de escrita.
  return (xml.match(/<m:oMath(?:Para)?[ >][\s\S]*?<\/m:oMath(?:Para)?>/g) ?? []).map((math) =>
    math.replaceAll(' />', '/>'),
  )
}

/**
 * Equações (M11, fase 1): desenhadas pelo MathML, só para leitura, e devolvidas ao
 * arquivo com o OMML como veio — também quando se edita o texto ao lado.
 */
test.describe('equações', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-equacoes-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('a equação em linha e a de exibição aparecem; editar o texto ao lado e reabrir as mantém', async () => {
    const origem = join(pasta, 'relatorio.docx')
    const bytes = await docxWithEquations()
    await writeFile(origem, bytes)
    const antes = equacoesDo(await entryOf(origem, 'word/document.xml'))
    expect(antes).toHaveLength(3)

    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    const equacoes = editor.locator('.equacao')
    await expect(equacoes).toHaveCount(3)

    // A em linha: dentro da frase, na altura do texto.
    const emLinha = equacoes.nth(0)
    await expect(emLinha.locator('math msup')).toHaveCount(1)
    const caixaEmLinha = await emLinha.boundingBox()
    expect(caixaEmLinha!.width).toBeGreaterThan(10)
    expect(caixaEmLinha!.height).toBeLessThan(40)

    // A de exibição: um bloco da largura da coluna, com a fração e a raiz.
    const exibicao = equacoes.nth(1)
    await expect(exibicao).toHaveClass(/equacao--exibicao/)
    await expect(exibicao.locator('math mfrac msqrt')).toHaveCount(1)
    const caixaExibicao = await exibicao.boundingBox()
    const caixaEditor = await editor.boundingBox()
    expect(caixaExibicao!.width).toBeGreaterThan(caixaEditor!.width * 0.8)
    expect(caixaExibicao!.height).toBeGreaterThan(caixaEmLinha!.height)

    // A travada diz o que a tela não desenha.
    const travada = equacoes.nth(2)
    await expect(travada).toHaveClass(/equacao--travada/)
    await expect(travada).toHaveAttribute('title', /m:borderBox/)

    await editor.getByText(EQUATION_AFTER.trim()).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Conferido.')
    await expect(editor).toContainText('para todo raio. Conferido.')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('para todo raio. Conferido.')
    expect(corpo).toContain(EQUATION_BEFORE)
    expect(equacoesDo(corpo)).toEqual(antes)

    await menu(session, 'open')
    await expect(editor).toContainText('para todo raio. Conferido.')
    await expect(editor.locator('.equacao')).toHaveCount(3)
    await expect(editor.locator('.equacao').nth(0).locator('math msup')).toHaveCount(1)
  })

  test('o PDF leva as equações', async () => {
    test.skip(!(await temPdftotext()), 'pdftotext não instalado')
    const origem = join(pasta, 'relatorio.docx')
    const destino = join(pasta, 'relatorio.pdf')
    await writeFile(origem, await docxWithEquations())
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .equacao')).toHaveCount(3)

    await menu(session, 'export-pdf')
    await expect.poll(() => textoDoPdf(destino), { timeout: 30_000 }).toContain('para todo raio')
    const texto = await textoDoPdf(destino)
    // O π da área, o Δ e a raiz de Bhaskara vêm das equações — o texto em volta
    // não os tem. O π e o Δ saem no itálico matemático (U+1D70B, U+1D6E5), que é
    // como o `mi` de uma letra só os desenha.
    expect(texto).toMatch(/[π𝜋]/u)
    expect(texto).toMatch(/[Δ𝛥]/u)
    expect(texto).toContain('√')
    expect(texto).toContain('Fim.')
  })

  // --- fase 2: o editor de equações ------------------------------------------

  test('inserir pelo menu, digitar o LaTeX, desfazer e refazer; salvar em .docx e reabrir', async () => {
    const destino = join(pasta, 'nova.docx')
    await stubDialogs(session.app, { save: destino, open: destino, messageBox: 1 })
    await menu(session, 'new-document')
    const editor = session.window.locator('.pages__column .ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Área: ')
    // Passado o intervalo do histórico, o texto é um passo e a equação é outro.
    await session.window.waitForTimeout(700)

    await menu(session, 'insert-equation')
    const dialogo = session.window.getByRole('dialog', { name: 'Equação', exact: true })
    const fonte = dialogo.locator('textarea')
    await expect(fonte).toBeFocused()
    await fonte.fill('\\frac{a}{b}+\\sqrt{x}')
    await expect(dialogo.locator('.equation__preview math mfrac')).toHaveCount(1)
    await dialogo.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialogo).toHaveCount(0)

    const equacao = editor.locator('.equacao')
    await expect(equacao).toHaveCount(1)
    await expect(equacao.locator('math mfrac')).toHaveCount(1)
    await expect(equacao.locator('math msqrt')).toHaveCount(1)

    // Um passo de desfazer por OK.
    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('.equacao')).toHaveCount(0)
    await expect(editor).toContainText('Área:')
    await session.window.keyboard.press('Control+y')
    await expect(editor.locator('.equacao math mfrac')).toHaveCount(1)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/document.xml')
    expect(corpo).toContain('<m:f>')
    expect(corpo).toContain('<m:rad>')
    expect(corpo).toContain('Cambria Math')

    await menu(session, 'close-file')
    await menu(session, 'open')
    await expect(editor.locator('.equacao')).toHaveCount(1)
    await expect(editor.locator('.equacao math mfrac')).toHaveCount(1)
    await expect(editor.locator('.equacao math msqrt')).toHaveCount(1)
  })

  test('editar uma equação do arquivo; cancelar não muda nada; a travada só abre para ver', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithEquations())
    const antes = equacoesDo(await entryOf(origem, 'word/document.xml'))
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    const equacoes = editor.locator('.equacao')
    await expect(equacoes).toHaveCount(3)
    const estado = session.window.locator('.statusbar__state')
    const estadoAntes = await estado.textContent()
    const dialogo = session.window.getByRole('dialog', { name: 'Equação', exact: true })

    // O LaTeX sai do MathML da equação do arquivo; cancelar deixa tudo como estava.
    await equacoes.nth(0).dblclick()
    await expect(dialogo.locator('textarea')).toHaveValue(/\\pi\s*r\^\{2\}/)
    await dialogo.getByRole('button', { name: 'Cancelar' }).click()
    await expect(dialogo).toHaveCount(0)
    await expect(estado).toHaveText(estadoAntes ?? '')

    // Selecionada, o Enter também abre.
    await equacoes.nth(0).click()
    await session.window.keyboard.press('Enter')
    await expect(dialogo).toHaveCount(1)
    await dialogo.locator('textarea').fill('\\pi r^{3}')
    await dialogo.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(equacoes.nth(0).locator('math msup mn')).toHaveText('3')
    await expect(equacoes).toHaveCount(3)

    // A travada lista só a construção que falta, e não abre para editar.
    await equacoes.nth(2).dblclick()
    await expect(dialogo).toContainText('m:borderBox')
    await expect(dialogo).not.toContainText('m:e')
    await expect(dialogo.locator('textarea')).toHaveCount(0)
    await dialogo.getByRole('button', { name: 'Fechar' }).click()

    await menu(session, 'save')
    await expect(estado).toHaveText('Salvo')
    const depois = equacoesDo(await entryOf(origem, 'word/document.xml'))
    expect(depois).toHaveLength(3)
    expect(depois[0]).toContain('<m:sSup>')
    expect(depois[0]).toMatch(/<m:t>3<\/m:t>/)
    // As outras duas voltam como vieram.
    expect(depois.slice(1)).toEqual(antes.slice(1))

    await menu(session, 'open')
    await expect(editor.locator('.equacao').nth(0).locator('math msup mn')).toHaveText('3')
  })

  // --- fase 3: exportações e integração ---------------------------------------

  test('exportar em HTML, Markdown e ODT leva as equações', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithEquations())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .equacao')).toHaveCount(3)
    const lido = (caminho: string) => () => readFile(caminho, 'utf8').catch(() => '')

    const html = join(pasta, 'relatorio.html')
    await stubDialogs(session.app, { save: html })
    await menu(session, 'export-html')
    await expect.poll(lido(html), { timeout: 15_000 }).toContain('</html>')
    const pagina = await readFile(html, 'utf8')
    expect(pagina).toMatch(/<math xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="inline">/)
    expect(pagina).toMatch(
      /<math xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="block">.*<mfrac>/,
    )

    const markdown = join(pasta, 'relatorio.md')
    await stubDialogs(session.app, { save: markdown })
    await menu(session, 'export-markdown')
    await expect.poll(lido(markdown), { timeout: 15_000 }).toContain(EQUATION_AFTER.trim())
    const texto = await readFile(markdown, 'utf8')
    // O LaTeX sai do MathML das equações do arquivo, que não o guardam.
    expect(texto).toMatch(/A área do círculo é \$\\pi\s*r\^\{2\}\$ para todo raio\./)
    expect(texto).toMatch(/^\$\$x\s*=.*\\frac\{.*\\Delta.*\$\$$/m)

    const odt = join(pasta, 'relatorio.odt')
    await stubDialogs(session.app, { save: odt })
    await menu(session, 'export-odt')
    await expect
      .poll(() => entryOf(odt, 'META-INF/manifest.xml').catch(() => ''), { timeout: 15_000 })
      .toContain('application/vnd.oasis.opendocument.formula')
    const conteudo = await entryOf(odt, 'content.xml')
    expect(conteudo.match(/<draw:object xlink:href="\.\/Object \d+"/g)).toHaveLength(3)
    expect(await entryOf(odt, 'Object 2/content.xml')).toMatch(/<math [^>]*display="block">.*<mfrac>/)
  })

  test('copiar uma equação dá o LaTeX; colar a nossa mantém o OMML, e o MathML de fora vira equação nova', async () => {
    const origem = join(pasta, 'relatorio.docx')
    await writeFile(origem, await docxWithEquations())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.pages__column .ProseMirror')
    const equacoes = editor.locator('.equacao')
    await expect(equacoes).toHaveCount(3)
    const omml = await equacoes.nth(0).getAttribute('data-omml')
    expect(omml).toContain('oMath')

    await equacoes.nth(0).click()
    await session.window.keyboard.press('ControlOrMeta+c')
    await expect
      .poll(() => session.app.evaluate(({ clipboard }) => clipboard.readText()))
      .toMatch(/\\pi\s*r\^\{2\}/)

    // Colada no fim do documento, é a mesma equação, com o OMML que veio do arquivo.
    await session.window.keyboard.press('ControlOrMeta+End')
    await session.window.keyboard.press('ControlOrMeta+v')
    await expect(equacoes).toHaveCount(4)
    await expect(equacoes.nth(3)).toHaveAttribute('data-omml', omml ?? '')

    // O MathML de uma página da Web: equação nova, sem OMML, com o LaTeX tirado dele.
    await session.app.evaluate(({ clipboard }) => {
      clipboard.write({
        text: 'a/b',
        html: '<p>Da Web: <math display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math></p>',
      })
    })
    await session.window.keyboard.press('ControlOrMeta+v')
    await expect(equacoes).toHaveCount(5)
    const colada = equacoes.nth(4)
    await expect(colada).toHaveAttribute('data-omml', '')
    await expect(colada).toHaveAttribute('data-latex', '\\frac{a}{b}')
    await expect(colada).toHaveAttribute('data-display', 'true')
    await expect(colada.locator('math mfrac')).toHaveCount(1)
  })
})
