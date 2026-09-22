import { describe, expect, it } from 'vitest'
import { MenuCommand } from '@shared/types.js'
import { EditorCommand, asEditorCommand, runsWhileLocked } from './editor-commands.js'

describe('comandos do editor no somente leitura', () => {
  it('só a busca e a contagem rodam com o documento travado', () => {
    const allowed = Object.values(EditorCommand).filter(runsWhileLocked)
    expect(allowed.sort()).toEqual([EditorCommand.FindReplace, EditorCommand.WordCount].sort())
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
