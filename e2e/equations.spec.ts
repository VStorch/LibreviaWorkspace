import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { EQUATION_AFTER, EQUATION_BEFORE, docxWithEquations, entryOf } from './fixtures.js'
import { hasPdftotext, pdfText } from './external-tools.js'

/** The XML's `m:oMath`/`m:oMathPara`, in order: what must come back as it came. */
function equationsOf(xml: string): string[] {
  // The whole `word/document.xml` is reserialized by the SDK, which closes an empty element with "
  // />": the only difference in writing.
  return (xml.match(/<m:oMath(?:Para)?[ >][\s\S]*?<\/m:oMath(?:Para)?>/g) ?? []).map((math) =>
    math.replaceAll(' />', '/>'),
  )
}

/**
 * Equations: drawn from MathML, read-only, and returned to the file with the OMML as it came, also
 * when editing the text beside them.
 */
test.describe('equations', () => {
  let session: Session
  let directory: string

  test.beforeEach(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-equations-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(directory, { recursive: true, force: true })
  })

  test('the inline and display equations appear; editing the text beside them and reopening keeps them', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    const bytes = await docxWithEquations()
    await writeFile(sourceFile, bytes)
    const before = equationsOf(await entryOf(sourceFile, 'word/document.xml'))
    expect(before).toHaveLength(3)

    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    await expect(editor).toHaveAttribute('contenteditable', 'true')
    const equations = editor.locator('.equation')
    await expect(equations).toHaveCount(3)

    // The inline one: inside the sentence, at text height.
    const inline = equations.nth(0)
    await expect(inline.locator('math msup')).toHaveCount(1)
    const inlineBox = await inline.boundingBox()
    expect(inlineBox!.width).toBeGreaterThan(10)
    expect(inlineBox!.height).toBeLessThan(40)

    // The display one: a column-wide block, with the fraction and the root.
    const display = equations.nth(1)
    await expect(display).toHaveClass(/equation--display/)
    await expect(display.locator('math mfrac msqrt')).toHaveCount(1)
    const displayBox = await display.boundingBox()
    const editorBox = await editor.boundingBox()
    expect(displayBox!.width).toBeGreaterThan(editorBox!.width * 0.8)
    expect(displayBox!.height).toBeGreaterThan(inlineBox!.height)

    // A locked one says what the screen does not draw.
    const locked = equations.nth(2)
    await expect(locked).toHaveClass(/equation--locked/)
    await expect(locked).toHaveAttribute('title', /m:borderBox/)

    await editor.getByText(EQUATION_AFTER.trim()).click()
    await session.window.keyboard.press('End')
    await session.window.keyboard.type(' Conferido.')
    await expect(editor).toContainText('para todo raio. Conferido.')

    await menu(session, 'save')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    const body = await entryOf(sourceFile, 'word/document.xml')
    expect(body).toContain('para todo raio. Conferido.')
    expect(body).toContain(EQUATION_BEFORE)
    expect(equationsOf(body)).toEqual(before)

    await menu(session, 'open')
    await expect(editor).toContainText('para todo raio. Conferido.')
    await expect(editor.locator('.equation')).toHaveCount(3)
    await expect(editor.locator('.equation').nth(0).locator('math msup')).toHaveCount(1)
  })

  test('the PDF carries the equations', async () => {
    test.skip(!(await hasPdftotext()), 'pdftotext not installed')
    const sourceFile = join(directory, 'relatorio.docx')
    const destinationFile = join(directory, 'relatorio.pdf')
    await writeFile(sourceFile, await docxWithEquations())
    await stubDialogs(session.app, { open: sourceFile, save: destinationFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .equation')).toHaveCount(3)

    await menu(session, 'export-pdf')
    await expect
      .poll(() => pdfText(destinationFile, { layout: false }), { timeout: 30_000 })
      .toContain('para todo raio')
    const text = await pdfText(destinationFile, { layout: false })
    // The area's π, the Δ and the quadratic formula's root come from the equations; the surrounding
    // text lacks them. π and Δ come out in mathematical italic (U+1D70B, U+1D6E5), which is how a
    // one-letter `mi` draws them.
    expect(text).toMatch(/[π𝜋]/u)
    expect(text).toMatch(/[Δ𝛥]/u)
    expect(text).toContain('√')
    expect(text).toContain('Fim.')
  })

  test('insert from the menu, type the LaTeX, undo and redo; save to .docx and reopen', async () => {
    const destinationFile = join(directory, 'nova.docx')
    await stubDialogs(session.app, { save: destinationFile, open: destinationFile, messageBox: 1 })
    await menu(session, 'new-document')
    const editor = session.window.locator('.pages__column .ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Área: ')
    // Past the history interval, the text is one step and the equation another.
    await session.window.waitForTimeout(700)

    await menu(session, 'insert-equation')
    const dialog = session.window.getByRole('dialog', { name: 'Equação', exact: true })
    const font = dialog.locator('textarea')
    await expect(font).toBeFocused()
    await font.fill('\\frac{a}{b}+\\sqrt{x}')
    await expect(dialog.locator('.equation__preview math mfrac')).toHaveCount(1)
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)

    const equation = editor.locator('.equation')
    await expect(equation).toHaveCount(1)
    await expect(equation.locator('math mfrac')).toHaveCount(1)
    await expect(equation.locator('math msqrt')).toHaveCount(1)

    // One undo step per OK.
    await session.window.keyboard.press('Control+z')
    await expect(editor.locator('.equation')).toHaveCount(0)
    await expect(editor).toContainText('Área:')
    await session.window.keyboard.press('Control+y')
    await expect(editor.locator('.equation math mfrac')).toHaveCount(1)

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')
    const body = await entryOf(destinationFile, 'word/document.xml')
    expect(body).toContain('<m:f>')
    expect(body).toContain('<m:rad>')
    expect(body).toContain('Cambria Math')

    await menu(session, 'close-file')
    await menu(session, 'open')
    await expect(editor.locator('.equation')).toHaveCount(1)
    await expect(editor.locator('.equation math mfrac')).toHaveCount(1)
    await expect(editor.locator('.equation math msqrt')).toHaveCount(1)
  })

  test('edit a file equation; canceling changes nothing; the locked one only opens to view', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithEquations())
    const before = equationsOf(await entryOf(sourceFile, 'word/document.xml'))
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')

    const editor = session.window.locator('.pages__column .ProseMirror')
    const equations = editor.locator('.equation')
    await expect(equations).toHaveCount(3)
    const state = session.window.locator('.statusbar__state')
    const stateBefore = await state.textContent()
    const dialog = session.window.getByRole('dialog', { name: 'Equação', exact: true })

    // LaTeX comes from the file equation's MathML; canceling leaves everything as it was.
    await equations.nth(0).dblclick()
    await expect(dialog.locator('textarea')).toHaveValue(/\\pi\s*r\^\{2\}/)
    await dialog.getByRole('button', { name: 'Cancelar' }).click()
    await expect(dialog).toHaveCount(0)
    await expect(state).toHaveText(stateBefore ?? '')

    // Selected, Enter also opens it.
    await equations.nth(0).click()
    await session.window.keyboard.press('Enter')
    await expect(dialog).toHaveCount(1)
    await dialog.locator('textarea').fill('\\pi r^{3}')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(equations.nth(0).locator('math msup mn')).toHaveText('3')
    await expect(equations).toHaveCount(3)

    // A locked one lists only the missing construct, and does not open for editing.
    await equations.nth(2).dblclick()
    await expect(dialog).toContainText('m:borderBox')
    await expect(dialog).not.toContainText('m:e')
    await expect(dialog.locator('textarea')).toHaveCount(0)
    await dialog.getByRole('button', { name: 'Fechar' }).click()

    await menu(session, 'save')
    await expect(state).toHaveText('Salvo')
    const after = equationsOf(await entryOf(sourceFile, 'word/document.xml'))
    expect(after).toHaveLength(3)
    expect(after[0]).toContain('<m:sSup>')
    expect(after[0]).toMatch(/<m:t>3<\/m:t>/)
    // The other two come back as they came.
    expect(after.slice(1)).toEqual(before.slice(1))

    await menu(session, 'open')
    await expect(editor.locator('.equation').nth(0).locator('math msup mn')).toHaveText('3')
  })

  test('exporting to HTML, Markdown and ODT carries the equations', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithEquations())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.pages__column .equation')).toHaveCount(3)
    const read = (filePath: string) => () => readFile(filePath, 'utf8').catch(() => '')

    const html = join(directory, 'relatorio.html')
    await stubDialogs(session.app, { save: html })
    await menu(session, 'export-html')
    await expect.poll(read(html), { timeout: 15_000 }).toContain('</html>')
    const page = await readFile(html, 'utf8')
    expect(page).toMatch(/<math xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="inline">/)
    expect(page).toMatch(/<math xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="block">.*<mfrac>/)

    const markdown = join(directory, 'relatorio.md')
    await stubDialogs(session.app, { save: markdown })
    await menu(session, 'export-markdown')
    await expect.poll(read(markdown), { timeout: 15_000 }).toContain(EQUATION_AFTER.trim())
    const text = await readFile(markdown, 'utf8')
    // LaTeX comes from the MathML of the file's equations, which do not store it.
    expect(text).toMatch(/A área do círculo é \$\\pi\s*r\^\{2\}\$ para todo raio\./)
    expect(text).toMatch(/^\$\$x\s*=.*\\frac\{.*\\Delta.*\$\$$/m)

    const odt = join(directory, 'relatorio.odt')
    await stubDialogs(session.app, { save: odt })
    await menu(session, 'export-odt')
    await expect
      .poll(() => entryOf(odt, 'META-INF/manifest.xml').catch(() => ''), { timeout: 15_000 })
      .toContain('application/vnd.oasis.opendocument.formula')
    const contents = await entryOf(odt, 'content.xml')
    expect(contents.match(/<draw:object xlink:href="\.\/Object \d+"/g)).toHaveLength(3)
    expect(await entryOf(odt, 'Object 2/content.xml')).toMatch(/<math [^>]*display="block">.*<mfrac>/)
  })

  test('copying an equation gives the LaTeX; pasting ours keeps the OMML, and outside MathML becomes a new equation', async () => {
    const sourceFile = join(directory, 'relatorio.docx')
    await writeFile(sourceFile, await docxWithEquations())
    await stubDialogs(session.app, { open: sourceFile, messageBox: 1 })
    await menu(session, 'open')
    const editor = session.window.locator('.pages__column .ProseMirror')
    const equations = editor.locator('.equation')
    await expect(equations).toHaveCount(3)
    const omml = await equations.nth(0).getAttribute('data-omml')
    expect(omml).toContain('oMath')

    await equations.nth(0).click()
    await session.window.keyboard.press('ControlOrMeta+c')
    await expect
      .poll(() => session.app.evaluate(({ clipboard }) => clipboard.readText()))
      .toMatch(/\\pi\s*r\^\{2\}/)

    // Pasted at the end of the document, it is the same equation, with the OMML from the file.
    await session.window.keyboard.press('ControlOrMeta+End')
    await session.window.keyboard.press('ControlOrMeta+v')
    await expect(equations).toHaveCount(4)
    await expect(equations.nth(3)).toHaveAttribute('data-omml', omml ?? '')

    // MathML from a web page: a new equation, without OMML, with LaTeX derived from it.
    await session.app.evaluate(({ clipboard }) => {
      clipboard.write({
        text: 'a/b',
        html: '<p>Da Web: <math display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math></p>',
      })
    })
    await session.window.keyboard.press('ControlOrMeta+v')
    await expect(equations).toHaveCount(5)
    const pasted = equations.nth(4)
    await expect(pasted).toHaveAttribute('data-omml', '')
    await expect(pasted).toHaveAttribute('data-latex', '\\frac{a}{b}')
    await expect(pasted).toHaveAttribute('data-display', 'true')
    await expect(pasted.locator('math mfrac')).toHaveCount(1)
  })
})
