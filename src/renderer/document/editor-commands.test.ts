import { describe, expect, it } from 'vitest'
import { MenuCommand } from '@shared/types.js'
import { EditorCommand, asEditorCommand, runsWhileLocked } from './editor-commands.js'

describe('comandos do editor no somente leitura', () => {
  it('só a busca, a contagem, os marcadores, a navegação e abrir a equação rodam com o documento travado', () => {
    // The bookmark dialog opens for "Go to"; add and delete grey out inside it. The equation opens
    // for viewing: the dialog does not save while locked.
    const allowed = Object.values(EditorCommand).filter(runsWhileLocked)
    expect(allowed.sort()).toEqual(
      [
        EditorCommand.FindReplace,
        EditorCommand.WordCount,
        EditorCommand.InsertBookmark,
        EditorCommand.NextComment,
        EditorCommand.PreviousComment,
        EditorCommand.AuthorName,
        EditorCommand.NextChange,
        EditorCommand.PreviousChange,
        EditorCommand.EditEquation,
      ].sort(),
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
