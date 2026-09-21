import { stat } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, type Session } from './app.js'

/**
 * As ferramentas do dia a dia, pelo aplicativo montado.
 *
 * Tudo aqui atravessa os dois processos, e é justamente a costura que teste de
 * unidade nenhum cobre: o menu de contexto nasce de um evento do `webContents`,
 * colar sem formatação depende da área de transferência do sistema, e a
 * ortografia depende de um dicionário que precisa estar em disco **antes** de o
 * Chromium perguntar por ele.
 *
 * `new-document` antes de tudo: sem arquivo aberto a tela é a inicial, e não há
 * editor onde clicar.
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
    // O modo de falha que este teste cobre é o silencioso: sem o dicionário no
    // lugar certo, o corretor não marca nada e ninguém é avisado — e numa máquina
    // sem rede é isso que acontecia, porque o Chromium baixaria o arquivo.
    const corretor = await session.app.evaluate(({ session: electronSession }) => ({
      idiomas: electronSession.defaultSession.getSpellCheckerLanguages(),
      ligado: electronSession.defaultSession.isSpellCheckerEnabled(),
    }))

    expect(corretor.idiomas).toEqual(['pt-BR'])
    expect(corretor.ligado).toBe(true)

    // O arquivo que o Chromium procura antes de pensar em baixar, com o tamanho
    // do que viaja no instalador: se fosse o baixado, seria outro.
    const instalado = await stat(join(session.userData, 'Dictionaries', 'pt-BR-3-0.bdic'))
    const embutido = await stat(resolve('resources/dictionaries/pt-BR-3-0.bdic'))
    expect(instalado.size).toBe(embutido.size)

    // E o campo de texto pede verificação: a outra metade do recurso.
    await expect(session.window.locator('.ProseMirror')).toHaveAttribute('spellcheck', 'true')
  })

  test('desligar a verificação ortográfica desliga o corretor e o campo', async () => {
    // Pelo mesmo caminho que o item do menu nativo usa — o menu nativo em si
    // nenhum teste consegue clicar.
    // O tipo de `window.api` mora na declaração do renderer, que este projeto de
    // TypeScript não inclui: daí o elenco, e não uma segunda declaração global.
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

  // Marcar a palavra errada é trabalho do corretor do Chromium, e ele só o faz
  // quando o quadro tem foco de verdade: numa janela que o gerenciador de janelas
  // não trouxe para a frente — o caso de toda execução automatizada — a marcação
  // não acontece de forma previsível, e o menu abre sem as sugestões. O caminho
  // até elas está coberto pelos testes acima (dicionário no lugar, corretor
  // ligado, campo pedindo verificação) e pelo menu de contexto, que é nosso.
  test.skip('a palavra errada traz sugestão, dicionário e ignorar', async () => {
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    // Em negrito para dar ao teste **onde clicar**: o parágrafo ocupa a coluna
    // inteira e o centro dele cairia longe da palavra, alinhada à esquerda.
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

    // O motivo de todo o cuidado com o CSS das marcas: se elas tivessem altura, a
    // linha cresceria e a quebra de página se deslocaria.
    await expect(paginas).toHaveText(antes ?? '')

    await session.window.getByRole('button', { name: /Marcas de formatação/ }).click()
    await expect(editor.locator('.tiptap-invisible-character')).toHaveCount(0)
  })

  test('o seletor de caracteres especiais também anda pelo teclado', async () => {
    // O defeito que este teste cobre: nada recebia foco ao abrir o painel, então as
    // setas moviam o cursor do texto, `Enter` partia o parágrafo e só o mouse
    // inseria. Pior, o `Tab` saía do painel e caía no seletor "Estilo" da barra.
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

    // Inseriu e **continuou no painel**: o foco não voltou para o texto, senão a
    // seta seguinte andaria pelo documento.
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
    // A saída para texto técnico sem desligar nada: `--silent` volta a ser
    // `--silent`. Quem trata a tecla é o `Keymap` do Tiptap, e o que a mantém de pé
    // é o `undoable` que o embrulho da autocorreção preserva
    // (ver `editor-extensions.ts`).
    const editor = session.window.locator('.ProseMirror')
    await editor.click()
    await session.window.keyboard.type('rodar --')
    await expect(editor).toContainText('rodar —')

    // Logo depois da substituição, e não três palavras adiante: é assim no Word, e
    // é o que o plugin de regras de entrada guarda — uma correção, a última.
    await session.window.keyboard.press('Backspace')
    await session.window.keyboard.type('silent')
    await expect(editor).toContainText('rodar --silent')
  })
})
