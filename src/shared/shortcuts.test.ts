import { describe, expect, it } from 'vitest'
import {
  SHORTCUTS,
  ShortcutOwner,
  type Shortcut,
  acceleratorOf,
  canonicalKeyOf,
  editorKeyOf,
  shortcutHintOf,
} from './shortcuts.js'

/**
 * Quem pede a mesma tecla que outro.
 *
 * A colisão que interessa é a de donos diferentes: acelerador de menu é registrado
 * no main e intercepta a tecla antes de o renderer vê-la, então o atalho do editor
 * morre calado. Devolver a lista, em vez de apenas sim ou não, é o que faz a falha
 * dizer qual tecla e quais atalhos.
 */
function crossOwnerCollisions(table: Readonly<Record<string, Shortcut>>): string[] {
  const byKey = new Map<string, { id: string; owner: string }[]>()
  for (const [id, shortcut] of Object.entries(table)) {
    const key = canonicalKeyOf(shortcut.key)
    byKey.set(key, [...(byKey.get(key) ?? []), { id, owner: shortcut.owner }])
  }

  return [...byKey.entries()]
    .filter(([, entries]) => new Set(entries.map((entry) => entry.owner)).size > 1)
    .map(([key, entries]) => `${key}: ${entries.map((entry) => `${entry.id} (${entry.owner})`).join(' e ')}`)
}

describe('tabela de atalhos', () => {
  it('não declara a mesma tecla para donos diferentes', () => {
    // A regressão que se quer impossível: um acelerador de menu novo em cima de
    // uma tecla do editor apaga o atalho do editor sem erro nenhum.
    expect(crossOwnerCollisions(SHORTCUTS)).toEqual([])
  })

  it('acusa a colisão quando ela existe — o teste acima não passa por vazio', () => {
    const collisions = crossOwnerCollisions({
      superscript: SHORTCUTS.superscript,
      // O "Ampliar" de antes, no `CommandOrControl+Plus` que o Electron traduz
      // para a tecla do `=` com Shift: a colisão que só apareceu pelo relato de
      // que o sobrescrito havia parado de funcionar.
      zoomIn: { owner: ShortcutOwner.Menu, key: { mod: true, shift: true, key: '=' }, does: 'Ampliar' },
    })

    expect(collisions).toEqual(['Mod+Shift+=: superscript (editor) e zoomIn (menu)'])
  })

  it('não declara a mesma tecla duas vezes, nem para o mesmo dono', () => {
    const keys = Object.values(SHORTCUTS).map((shortcut) => canonicalKeyOf(shortcut.key))
    expect(new Set(keys).size).toBe(keys.length)
  })

  it('escreve o acelerador do menu como o Electron o espera', () => {
    // Fixados um a um: estas são as teclas que o aplicativo tem hoje, e a tabela
    // não pode tê-las mudado ao juntá-las num lugar só.
    expect(acceleratorOf(SHORTCUTS.newDocument)).toBe('CmdOrCtrl+N')
    expect(acceleratorOf(SHORTCUTS.newSpreadsheet)).toBe('CmdOrCtrl+Shift+N')
    expect(acceleratorOf(SHORTCUTS.open)).toBe('CmdOrCtrl+O')
    expect(acceleratorOf(SHORTCUTS.save)).toBe('CmdOrCtrl+S')
    expect(acceleratorOf(SHORTCUTS.saveAs)).toBe('CmdOrCtrl+Shift+S')
    expect(acceleratorOf(SHORTCUTS.print)).toBe('CmdOrCtrl+P')
    expect(acceleratorOf(SHORTCUTS.closeFile)).toBe('CmdOrCtrl+W')
    expect(acceleratorOf(SHORTCUTS.pasteWithoutFormat)).toBe('CmdOrCtrl+Shift+V')
    expect(acceleratorOf(SHORTCUTS.findReplace)).toBe('CmdOrCtrl+F')
    expect(acceleratorOf(SHORTCUTS.insertPageBreak)).toBe('CmdOrCtrl+Enter')
    expect(acceleratorOf(SHORTCUTS.formattingMarks)).toBe('CmdOrCtrl+F10')
    expect(acceleratorOf(SHORTCUTS.zoomIn)).toBe('CmdOrCtrl+numadd')
    expect(acceleratorOf(SHORTCUTS.reload)).toBe('CmdOrCtrl+Shift+R')
    expect(acceleratorOf(SHORTCUTS.wordCount)).toBe('CmdOrCtrl+Shift+G')
  })

  it('escreve a tecla do editor com a letra minúscula que o prosemirror-keymap exige', () => {
    // Com `L` maiúsculo o keymap entenderia "a tecla que só sai com Shift", e o
    // atalho nunca dispararia.
    expect(editorKeyOf(SHORTCUTS.alignLeft)).toBe('Mod-l')
    expect(editorKeyOf(SHORTCUTS.alignCenter)).toBe('Mod-e')
    expect(editorKeyOf(SHORTCUTS.alignRight)).toBe('Mod-r')
    expect(editorKeyOf(SHORTCUTS.alignJustify)).toBe('Mod-j')
    expect(editorKeyOf(SHORTCUTS.lineHeightSingle)).toBe('Mod-1')
    expect(editorKeyOf(SHORTCUTS.lineHeightOneAndHalf)).toBe('Mod-5')
    expect(editorKeyOf(SHORTCUTS.lineHeightDouble)).toBe('Mod-2')
    expect(editorKeyOf(SHORTCUTS.superscript)).toBe('Mod-Shift-=')
    expect(editorKeyOf(SHORTCUTS.subscript)).toBe('Mod-=')
  })

  it('anuncia na barra de ferramentas exatamente as dicas de sempre', () => {
    expect(shortcutHintOf(SHORTCUTS.bold)).toBe('Ctrl+B')
    expect(shortcutHintOf(SHORTCUTS.italic)).toBe('Ctrl+I')
    expect(shortcutHintOf(SHORTCUTS.underline)).toBe('Ctrl+U')
    expect(shortcutHintOf(SHORTCUTS.superscript)).toBe('Ctrl+Shift+=')
    expect(shortcutHintOf(SHORTCUTS.subscript)).toBe('Ctrl+=')
    expect(shortcutHintOf(SHORTCUTS.alignLeft)).toBe('Ctrl+L')
    expect(shortcutHintOf(SHORTCUTS.alignCenter)).toBe('Ctrl+E')
    expect(shortcutHintOf(SHORTCUTS.alignRight)).toBe('Ctrl+R')
    expect(shortcutHintOf(SHORTCUTS.alignJustify)).toBe('Ctrl+J')
    expect(shortcutHintOf(SHORTCUTS.outdent)).toBe('Ctrl+[')
    expect(shortcutHintOf(SHORTCUTS.indent)).toBe('Ctrl+]')
    expect(shortcutHintOf(SHORTCUTS.insertPageBreak)).toBe('Ctrl+Enter')
    expect(shortcutHintOf(SHORTCUTS.formattingMarks)).toBe('Ctrl+F10')
    expect(shortcutHintOf(SHORTCUTS.findReplace)).toBe('Ctrl+F')
  })
})
