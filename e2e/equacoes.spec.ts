import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
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
})
