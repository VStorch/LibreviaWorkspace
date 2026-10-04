import { stat } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * As ferramentas do dia a dia pelo aplicativo montado, na costura entre os dois
 * processos: o menu de contexto do `webContents`, a área de transferência do
 * sistema e o dicionário que tem de estar em disco antes de o Chromium perguntar.
 * `new-document` primeiro, porque a tela inicial não tem editor.
 */
test.describe('ferramentas do documento', () => {
  let session: Session

  test.beforeEach(async () => {
    session = await launch()
    await menu(session, 'new-document')
  })

  test.afterEach(async () => {
    await session.close()
  })

  test('o botão direito abre o menu com recortar, copiar e colar', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Texto para o menu.')

    await editor.click({ button: 'right' })

    const contexto = session.window.getByRole('menu', { name: 'Ações do documento' })
    await expect(contexto).toBeVisible()
    await expect(contexto.getByRole('menuitem', { name: 'Copiar' })).toBeVisible()
    await expect(contexto.getByRole('menuitem', { name: 'Colar sem formatação' })).toBeVisible()

    // Escape fecha, como todo menu de contexto do sistema.
    await session.window.keyboard.press('Escape')
    await expect(contexto).toBeHidden()
  })

  test('o dicionário de português está instalado e o corretor, ligado', async () => {
    // Sem o dicionário no lugar, o corretor não marca nada, e ninguém é avisado.
    const corretor = await session.app.evaluate(({ session: electronSession }) => ({
      idiomas: electronSession.defaultSession.getSpellCheckerLanguages(),
      ligado: electronSession.defaultSession.isSpellCheckerEnabled(),
    }))

    expect(corretor.idiomas).toEqual(['pt-BR'])
    expect(corretor.ligado).toBe(true)

    // O arquivo que o Chromium procura antes de baixar, com o tamanho do que viaja no instalador.
    const instalado = await stat(join(session.userData, 'Dictionaries', 'pt-BR-3-0.bdic'))
    const embutido = await stat(resolve('resources/dictionaries/pt-BR-3-0.bdic'))
    expect(instalado.size).toBe(embutido.size)

    // E o campo de texto pede verificação: a outra metade do recurso.
    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('spellcheck', 'true')
  })

  test('desligar a verificação ortográfica desliga o corretor e o campo', async () => {
    // Pelo caminho do item do menu nativo, que teste nenhum clica. O tipo de
    // `window.api` mora na declaração do renderer, fora deste projeto: daí o elenco.
    await session.window.evaluate(async () => {
      const api = (
        window as unknown as { api: { preferences: { set: (patch: unknown) => Promise<unknown> } } }
      ).api
      await api.preferences.set({ spellcheck: false })
    })

    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('spellcheck', 'false')
    expect(
      await session.app.evaluate(({ session: electronSession }) =>
        electronSession.defaultSession.isSpellCheckerEnabled(),
      ),
    ).toBe(false)
  })

  // O Chromium só marca a palavra errada com o quadro em foco de verdade, o que a
  // execução automatizada não garante. O caminho até as sugestões está coberto acima.
  test.skip('a palavra errada traz sugestão, dicionário e ignorar', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    // Em negrito para ter onde clicar: o centro do parágrafo cairia longe da palavra.
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type('abacaxxi')
    await session.window.keyboard.press('Control+b')
    await session.window.keyboard.type(' e mais texto.')

    const contexto = session.window.getByRole('menu', { name: 'Ações do documento' })

    await expect(async () => {
      await session.window.keyboard.press('Escape')
      await editor.locator('strong').click({ button: 'right' })
      await expect(contexto.getByRole('menuitem', { name: 'Adicionar ao dicionário' })).toBeVisible({
        timeout: 1500,
      })
    }).toPass({ timeout: 20_000 })

    await expect(contexto.getByRole('menuitem', { name: 'Ignorar nesta sessão' })).toBeVisible()
  })

  test('colar sem formatação traz o texto e deixa a formatação de fora', async () => {
    await session.app.evaluate(({ clipboard }) => {
      clipboard.write({
        text: 'primeira linha\nsegunda linha',
        html: '<p><strong>primeira linha</strong></p><p><em>segunda linha</em></p>',
      })
    })

    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await menu(session, 'paste-without-format')

    await expect(editor).toContainText('primeira linha')
    await expect(editor).toContainText('segunda linha')
    // Duas linhas viram dois parágrafos, e nenhum deles traz o negrito nem o
    // itálico que vinham no HTML da área de transferência.
    await expect(editor.locator('strong')).toHaveCount(0)
    await expect(editor.locator('em')).toHaveCount(0)
  })

  test('a contagem de palavras conta o documento e a seleção', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('uma frase de teste')

    await menu(session, 'word-count')

    const dialogo = session.window.getByRole('dialog', { name: 'Contagem de palavras' })
    await expect(dialogo).toBeVisible()

    const palavras = dialogo.getByRole('row').filter({ hasText: 'Palavras' })
    await expect(palavras).toContainText('4')

    // Sem seleção a coluna da direita é um travessão: "nada selecionado" e
    // "seleção de zero palavras" não são a mesma coisa.
    await expect(palavras).toContainText('—')
  })

  test('o seletor de caracteres especiais insere no cursor', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Bom dia')

    await menu(session, 'special-character')

    const dialogo = session.window.getByRole('dialog', { name: 'Caracteres especiais' })
    await expect(dialogo).toBeVisible()

    // Pelo nome do caractere, que é o rótulo acessível do botão: apontar pelo
    // glifo dependeria de o teste conseguir digitá-lo.
    await dialogo.getByRole('button', { name: 'travessão' }).click()
    await expect(editor).toContainText('Bom dia—')
  })

  test('as marcas de formatação aparecem sem mexer na paginação', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Uma linha com espaços.')

    const paginas = session.window.locator('.statusbar__metric').filter({ hasText: 'página' })
    const antes = await paginas.textContent()

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()

    // Uma marca por espaço e uma no fim do parágrafo.
    await expect(editor.locator('.tiptap-invisible-character')).not.toHaveCount(0)

    // Marca com altura cresceria a linha e deslocaria a quebra de página.
    await expect(paginas).toHaveText(antes ?? '')

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()
    await expect(editor.locator('.tiptap-invisible-character')).toHaveCount(0)
  })

  test('o seletor de caracteres especiais também anda pelo teclado', async () => {
    // Sem foco no painel, as setas e o `Enter` iriam ao texto, e o `Tab` ao seletor "Estilo".
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Bom dia')

    await menu(session, 'special-character')
    const dialogo = session.window.getByRole('dialog', { name: 'Caracteres especiais' })

    // O primeiro caractere da grade já está com o foco; a seta anda, o Enter insere.
    const primeiro = dialogo.getByRole('button').first()
    await expect(primeiro).toBeFocused()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Enter')

    // Inseriu e continuou no painel.
    await expect(dialogo).toBeVisible()
    await expect(editor).not.toHaveText('Bom dia')
    const depoisDoPrimeiro = await editor.textContent()

    await session.window.keyboard.press('ArrowRight')
    await session.window.keyboard.press('Enter')
    await expect(editor).not.toHaveText(depoisDoPrimeiro ?? '')

    // O Tab circula dentro do painel: o "Estilo" da barra não é alcançado daqui.
    await session.window.keyboard.press('Tab')
    await expect(session.window.getByRole('combobox', { name: 'Estilo' })).not.toBeFocused()

    await session.window.keyboard.press('Escape')
    await expect(dialogo).toBeHidden()
  })

  test('a autocorreção tipográfica troca as aspas e o travessão', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('Ele disse "sim" -- e saiu')

    // Aspas curvas e travessão, como no Word em português.
    await expect(editor).toContainText('“sim”')
    await expect(editor).toContainText('—')
  })

  test('o Backspace desfaz só a substituição da autocorreção', async () => {
    // `--silent` volta a ser `--silent`: o `Keymap` do Tiptap trata a tecla, e o
    // `undoable` que o embrulho da autocorreção preserva a mantém (`editor-extensions.ts`).
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('rodar --')
    await expect(editor).toContainText('rodar —')

    // Logo depois da substituição, como no Word: o plugin guarda só a última correção.
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type('silent')
    await expect(editor).toContainText('rodar --silent')
  })
})
