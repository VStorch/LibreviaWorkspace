import { execFile } from 'node:child_process'
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithFootnote, docxWithLongFootnote, docxWithManyFootnotes, entryOf } from './fixtures.js'
import { hasPdfinfo, hasSoffice } from './external-tools.js'

/**
 * Footnotes and endnotes: read, numbered and preserved.
 *
 * The reference is a node with the note body inside. Editing the paragraph carrying it does not
 * lose it, which is why a document with notes does not open locked.
 */
test.describe('notas de rodapé', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-notas-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  test('editar o parágrafo da referência, salvar e reabrir mantém a nota', async () => {
    const origem = join(pasta, 'ata.docx')
    await writeFile(origem, await docxWithFootnote())
    await stubDialogs(session.app, { open: origem, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    // The note shows at the foot of the first sheet, with its number.
    const nota = session.window.locator('.paper-notes .note-body')
    await expect(nota).toContainText('Fonte: ata anterior.')
    await expect(nota.locator('.note-number')).toHaveText('1')

    await editor.getByText('Segundo parágrafo').click()
    await session.window.keyboard.press('Home')
    await session.window.keyboard.type('Revisto: ')
    await expect(editor).toContainText('Revisto: Segundo parágrafo')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // The paragraph was rewritten with the reference; the note, which nobody touched, stays in the
    // file, with the separators.
    const corpo = await entryOf(origem, 'word/document.xml')
    expect(corpo).toContain('Revisto: Segundo parágrafo')
    expect(corpo).toMatch(/<w:footnoteReference w:id="1" ?\/>/)
    const notas = await entryOf(origem, 'word/footnotes.xml')
    expect(notas).toContain('Fonte: ata anterior.')
    expect(notas).toContain('w:separator')

    await menu(session, 'open')
    await expect(editor).toContainText('Revisto: Segundo parágrafo')
    await expect(editor.locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    await expect(session.window.locator('.paper-notes .note-body')).toContainText('Fonte: ata anterior.')
    await expect(session.window.locator('.banner--notice')).toBeHidden()
  })
})

/**
 * Footnotes and endnotes: the body is edited at the foot of the page, and notes take room on the
 * sheet, so screen and paper break at the same place.
 */
test.describe('notas no pé da página', () => {
  let session: Session
  let pasta: string

  test.beforeEach(async () => {
    pasta = await mkdtemp(join(tmpdir(), 'librevia-notas-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(pasta, { recursive: true, force: true })
  })

  const texto = (session: Session) => session.window.locator('.pages__column .ProseMirror')
  const notas = (session: Session) => session.window.locator('.paper-notes .note-body')

  test('inserir nota pelo comando, escrever nela, desfazer e refazer', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Ata da reunião')
    await menu(session, 'insert-footnote')

    await expect(texto(session).locator('sup.note-ref')).toHaveAttribute('data-note-number', '1')
    await expect(notas(session)).toHaveCount(1)
    // The cursor is already in the note body.
    await expect(notas(session)).toBeFocused()
    await session.window.keyboard.type('Fonte: livro de atas.')
    await expect(notas(session)).toContainText('Fonte: livro de atas.')
    await expect(notas(session).locator('.note-number')).toHaveText('1')

    // A single history: undoing inside the note undoes in the document.
    await session.window.keyboard.press('Control+z')
    await expect(notas(session)).not.toContainText('Fonte: livro de atas.')
    await session.window.keyboard.press('Control+y')
    await expect(notas(session)).toContainText('Fonte: livro de atas.')

    // Clicking the number goes back to the reference; the next one, at the end, is 2.
    await notas(session).locator('.note-number').click()
    await expect(texto(session)).toBeFocused()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' e mais')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Segunda nota.')
    await expect(notas(session)).toHaveCount(2)
    await expect(notas(session).nth(1)).toContainText('Segunda nota.')
    await expect(notas(session).nth(1).locator('.note-number')).toHaveText('2')
  })

  test('apagar a referência apaga a nota; salvar em .docx e reabrir mantém o texto', async () => {
    const destino = join(pasta, 'ata.docx')
    await stubDialogs(session.app, { save: destino, open: destino, messageBox: 1 })
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Primeira')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota que fica.')
    await notas(session).locator('.note-number').click()
    await session.window.keyboard.type(' segunda')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota que sai.')
    await expect(notas(session)).toHaveCount(2)

    // Backspace right after the reference deletes the reference and the note.
    await notas(session).nth(1).locator('.note-number').click()
    await session.window.keyboard.press('Backspace')
    await expect(texto(session).locator('sup.note-ref')).toHaveCount(1)
    await expect(notas(session)).toHaveCount(1)
    await expect(notas(session)).toContainText('Nota que fica.')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const corpo = await entryOf(destino, 'word/footnotes.xml')
    expect(corpo).toContain('Nota que fica.')
    expect(corpo).not.toContain('Nota que sai.')

    await menu(session, 'close-file')
    await menu(session, 'open')
    await expect(notas(session)).toHaveCount(1)
    await expect(notas(session)).toContainText('Nota que fica.')
  })

  test('as notas empurram linhas, e a tela e o PDF têm as mesmas folhas', async () => {
    test.skip(!(await hasPdfinfo()), 'pdfinfo não instalado')
    const origem = join(pasta, 'longo.docx')
    const pdf = join(pasta, 'longo.pdf')
    await writeFile(origem, await docxWithManyFootnotes())
    await stubDialogs(session.app, { open: origem, save: pdf, messageBox: 1 })
    await menu(session, 'open')
    await expect(texto(session)).toContainText('24. Os conselheiros')

    const folhas = session.window.locator('.paper')
    await expect(notas(session)).toHaveCount(24)
    const naTela = await folhas.count()
    expect(naTela).toBeGreaterThan(2)
    // Every sheet with a reference has the notes area, at the foot.
    await expect(session.window.locator('.paper-notes--footnote')).toHaveCount(naTela)

    await menu(session, 'export-pdf')
    await expect.poll(() => paginasDoPdf(pdf), { timeout: 30_000 }).toBe(naTela)
    // Note 1 comes out on paper, on the first sheet.
    const { stdout } = await promisify(execFile)('pdftotext', ['-f', '1', '-l', '1', pdf, '-'])
    expect(stdout).toContain('Nota 1.')

    // And LibreOffice, with the same file, reaches the same sheet count.
    if (await hasSoffice()) {
      const copia = join(pasta, 'lo.docx')
      await copyFile(origem, copia)
      await promisify(execFile)('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', pasta, copia], {
        timeout: 120_000,
      })
      expect(await paginasDoPdf(join(pasta, 'lo.pdf'))).toBe(naTela)
    }
  })

  test('com o controle de alterações ligado, o que se digita na nota fica marcado', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Texto')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Original.')
    await expect(notas(session)).toContainText('Original.')

    await menu(session, 'toggle-track-changes')
    await notas(session).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Novo')
    await expect(notas(session).locator('ins')).toHaveText(' Novo')
    // A tracked Backspace leaves the deleted text in place, struck through.
    await session.window.keyboard.press('End')
    for (let vez = 0; vez < ' Novo'.length; vez++) await session.window.keyboard.press('ArrowLeft')
    await session.window.keyboard.press('Backspace')
    await expect(notas(session).locator('del')).toHaveText('.')
    await expect(notas(session)).toContainText('Original. Novo')
  })

  test('a nota de fim vai depois do último parágrafo', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Corpo do texto')
    await menu(session, 'insert-endnote')
    await session.window.keyboard.type('Ao fim do documento.')
    await expect(texto(session).locator('sup.note-ref')).toHaveAttribute('data-note-number', 'i')

    const area = session.window.locator('.paper-notes--endnote')
    await expect(area).toHaveCount(1)
    await expect(area.locator('.note-body')).toContainText('Ao fim do documento.')
    await expect(area.locator('.note-number')).toHaveText('i')
    // Right below the text, not at the foot of the sheet.
    const paragrafo = await texto(session).locator('p').first().boundingBox()
    const caixa = await area.boundingBox()
    expect(caixa!.y - (paragrafo!.y + paragrafo!.height)).toBeLessThan(60)

    // Search finds the note text.
    await menu(session, 'find-replace')
    await session.window.keyboard.type('fim do documento')
    await expect(area.locator('.search-hit')).toHaveCount(1)
  })

  test('com vários parágrafos, a nota de fim fica abaixo do último, ao reabrir também', async () => {
    const destino = join(pasta, 'fim.docx')
    await stubDialogs(session.app, { save: destino, open: destino, messageBox: 1 })
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Primeiro parágrafo')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Segundo parágrafo')
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Terceiro')
    await menu(session, 'insert-endnote')
    await session.window.keyboard.type('Fim novo.')

    const abaixoDoTexto = async (): Promise<void> => {
      const area = session.window.locator('.paper-notes--endnote')
      await expect(area).toHaveCount(1)
      const ultimo = await texto(session).locator('p').last().boundingBox()
      const caixa = await area.boundingBox()
      expect(caixa!.y).toBeGreaterThanOrEqual(ultimo!.y + ultimo!.height - 1)
    }
    await abaixoDoTexto()

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    await menu(session, 'close-file')
    await menu(session, 'open')
    await expect(texto(session)).toContainText('Terceiro')
    await abaixoDoTexto()
  })

  test('a nota maior que a folha continua na seguinte, sem cobrir o texto', async () => {
    const origem = join(pasta, 'nota-longa.docx')
    const pdf = join(pasta, 'nota-longa.pdf')
    await writeFile(origem, await docxWithLongFootnote())
    await stubDialogs(session.app, { open: origem, save: pdf, messageBox: 1 })
    await menu(session, 'open')
    await expect(texto(session)).toContainText('Delta')
    await expect(session.window.locator('.paper-notes__slot--continued')).toHaveCount(1)

    const geometria = await session.window.evaluate(() => {
      const caixa = (elemento: Element): { top: number; bottom: number } => {
        const { top, bottom } = elemento.getBoundingClientRect()
        return { top, bottom }
      }
      const folhas = Array.from(document.querySelectorAll('.paper'), caixa)
      const blocos = Array.from(document.querySelectorAll('.pages__column .ProseMirror > *'), caixa)
      const recortes = Array.from(document.querySelectorAll('.paper-notes__slot'), (slot) => {
        const area = caixa(slot)
        const visiveis = Array.from(slot.querySelectorAll('p')).filter((p) => {
          const linha = p.getBoundingClientRect()
          return linha.bottom > area.top + 1 && linha.top < area.bottom - 1
        })
        return {
          ...area,
          primeira: visiveis[0]?.textContent ?? '',
          ultima: visiveis.at(-1)?.textContent ?? '',
        }
      })
      return { folhas, blocos, recortes }
    })
    // Each slice within a sheet, without covering any text paragraph.
    for (const recorte of geometria.recortes) {
      expect(
        geometria.folhas.some((folha) => recorte.top >= folha.top && recorte.bottom <= folha.bottom),
      ).toBe(true)
      for (const bloco of geometria.blocos) {
        expect(recorte.bottom <= bloco.top + 1 || recorte.top >= bloco.bottom - 1).toBe(true)
      }
    }
    // The continuation starts where the first sheet stopped and goes to the end of the note.
    const [primeira, continuacao] = geometria.recortes
    const numero = (linha: string): number => Number(/Linha (\d+)/.exec(linha)?.[1] ?? -1)
    expect(numero(continuacao!.primeira)).toBe(numero(primeira!.ultima) + 1)
    expect(continuacao!.ultima).toContain('Linha 69')

    if (await hasPdfinfo()) {
      await menu(session, 'export-pdf')
      await expect.poll(() => paginasDoPdf(pdf), { timeout: 30_000 }).toBeGreaterThan(1)
      const { stdout } = await promisify(execFile)('pdftotext', ['-f', '2', '-l', '2', pdf, '-'])
      expect(stdout).toContain(`Linha ${numero(continuacao!.primeira)} `)
      expect(stdout).toContain('Linha 69')
      expect(stdout).not.toContain('Nota longa.')
    }
  })

  test('digitar depressa na nota até passar da folha não escapa para o texto', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Corpo')
    await menu(session, 'insert-footnote')
    for (let linha = 0; linha < 70; linha++) {
      await session.window.keyboard.type(`L${linha} nota longa que precisa continuar.`)
      await session.window.keyboard.press('Enter')
    }
    await session.window.keyboard.type('Fim da nota.')
    await expect(session.window.locator('.paper-notes__slot--continued').first()).toBeAttached()
    // The live body; continuations are copies of it.
    const nota = session.window.locator('.paper-notes__slot:not(.paper-notes__slot--continued) > .note-body')
    await expect(nota).toBeFocused()
    await expect(texto(session).locator('p').first()).toHaveText('Corpo')
    await expect(texto(session).locator('p')).toHaveCount(1)
    await expect(nota.locator('p')).toHaveCount(71)
    await expect(notas(session).locator('p').last()).toHaveText('Fim da nota.')
  })

  test('desfazer a nota com o cursor nela devolve o foco ao texto, e refazer funciona', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Texto')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota')
    await expect(notas(session)).toBeFocused()
    // The note text, then the note.
    await session.window.keyboard.press('Control+z')
    await session.window.keyboard.press('Control+z')
    await expect(notas(session)).toHaveCount(0)
    await expect(texto(session)).toBeFocused()

    await session.window.keyboard.press('Control+y')
    await expect(notas(session)).toHaveCount(1)
    await expect(texto(session).locator('sup.note-ref')).toHaveCount(1)
  })

  test('aceitar pelo menu de contexto a alteração no cursor dentro da nota', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Texto')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Original.')
    await expect(notas(session)).toContainText('Original.')

    await menu(session, 'toggle-track-changes')
    await notas(session).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Novo')
    await session.window.keyboard.type(' Outro')
    await expect(notas(session).locator('ins')).toHaveText(' Novo Outro')

    await notas(session).locator('ins').click({ button: 'right' })
    const contexto = session.window.getByRole('menu', { name: 'Ações do documento' })
    await contexto.getByRole('menuitem', { name: 'Aceitar alteração' }).click()
    await expect(notas(session).locator('ins')).toHaveCount(0)
    await expect(notas(session)).toContainText('Original. Novo Outro')
  })

  test('comentário novo dentro da nota é recusado, e o texto fica como estava', async () => {
    // LibreOffice does not open a .docx with `w:commentReference` in `footnotes.xml`.
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Ata')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Fonte da ata.')
    await session.window.keyboard.press('End')
    await session.window.keyboard.press('ArrowLeft')
    for (let vez = 0; vez < 'ata'.length; vez++) await session.window.keyboard.press('Shift+ArrowLeft')
    await menu(session, 'insert-comment')

    await expect(session.window.getByText(/não abriria o arquivo/)).toBeVisible()
    await expect(session.window.locator('.comment-card')).toHaveCount(0)
    await expect(notas(session).locator('.comment-range')).toHaveCount(0)
    await expect(notas(session)).toContainText('Fonte da ata.')
  })

  test('a referência cruzada a uma nota mostra o número e o F9 o atualiza', async () => {
    await menu(session, 'new-document')
    await texto(session).click()
    await session.window.keyboard.type('Primeira')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota alvo.')
    await notas(session).locator('.note-number').click()
    await session.window.keyboard.press('Enter')
    await session.window.keyboard.type('Ver nota ')

    await menu(session, 'insert-cross-reference')
    const ref = session.window.getByRole('dialog', { name: 'Referência cruzada' })
    await ref.getByLabel('Tipo').selectOption('note:footnote')
    await ref.getByLabel('Para qual').selectOption({ label: '1 Nota alvo.' })
    await expect(ref.getByLabel('Inserir referência a')).toHaveValue('number')
    await ref.getByRole('button', { name: 'Inserir' }).click()
    const paragrafo = texto(session).locator('p', { hasText: 'Ver nota' })
    await expect(paragrafo).toHaveText('Ver nota 1')

    // A new note before the cited one: F9 now cites 2.
    await texto(session).locator('p').first().click()
    await session.window.keyboard.press('Home')
    await menu(session, 'insert-footnote')
    await session.window.keyboard.type('Nota nova.')
    await expect(texto(session).locator('sup.note-ref').nth(1)).toHaveAttribute('data-note-number', '2')
    await menu(session, 'update-fields')
    await expect(paragrafo).toHaveText('Ver nota 2')
  })

  test('comentar uma palavra do parágrafo que tem nota ancora o comentário nela', async () => {
    const origem = join(pasta, 'comentario.docx')
    const destino = join(pasta, 'comentario-salvo.docx')
    await writeFile(origem, await docxWithManyFootnotes(2))
    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(notas(session)).toHaveCount(2)
    // Editing an earlier note, and going back to the text.
    await notas(session).first().click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Editada.')
    await texto(session).click()
    await session.window.keyboard.press('Control+Home')
    for (let vez = 0; vez < '1. '.length; vez++) await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Shift+ArrowRight')
    await session.window.keyboard.press('Shift+ArrowRight')
    await menu(session, 'insert-comment')
    await session.window.keyboard.type('Ver isto')
    await session.window.keyboard.press('Control+Enter')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const documento = await entryOf(destino, 'word/document.xml')
    const paragrafo = /<w:p>(?:(?!<w:p>).)*?<\/w:p>/s.exec(documento)?.[0] ?? ''
    expect(paragrafo).toMatch(
      /<w:commentRangeStart[^>]*\/>(<w:r>(?:(?!<\/w:r>).)*<\/w:r>)*?<w:commentRangeEnd/s,
    )
    const comentado = /<w:commentRangeStart[^>]*\/>(.*?)<w:commentRangeEnd/s.exec(paragrafo)?.[1] ?? ''
    expect(comentado.replace(/<[^>]+>/g, '')).toBe('Os')
    expect(paragrafo).toContain('<w:commentReference')
  })
})

async function paginasDoPdf(caminho: string): Promise<number> {
  try {
    const { stdout } = await promisify(execFile)('pdfinfo', [caminho])
    return Number(/Pages:\s+(\d+)/.exec(stdout)?.[1] ?? 0)
  } catch {
    return 0
  }
}
