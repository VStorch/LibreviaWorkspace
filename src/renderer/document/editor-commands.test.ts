import { describe, expect, it } from 'vitest'
import { MenuCommand } from '@shared/types.js'
import { EditorCommand, asEditorCommand, runsWhileLocked } from './editor-commands.js'

describe('comandos do editor no somente leitura', () => {
  it('só a busca, a contagem e os marcadores rodam com o documento travado', () => {
    // O diálogo de marcadores abre para "Ir para"; adicionar e excluir se apagam
    // lá dentro.
    const allowed = Object.values(EditorCommand).filter(runsWhileLocked)
    expect(allowed.sort()).toEqual(
      [EditorCommand.FindReplace, EditorCommand.WordCount, EditorCommand.InsertBookmark].sort(),
    )
  })

  it('toda ação de tabela é edição', () => {
    for (const command of Object.values(EditorCommand).filter((value) => value.startsWith('table-'))) {
      expect(runsWhileLocked(command)).toBe(false)
    }
  })
})

describe('repasse do menu ao editor', () => {
  it('reconhece pelo nome cada comando de editor, e só eles', () => {
    for (const command of Object.values(EditorCommand)) expect(asEditorCommand(command)).toBe(command)
    expect(asEditorCommand(MenuCommand.Save)).toBeNull()
  })
})
