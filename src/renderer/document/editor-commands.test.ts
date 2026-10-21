import { describe, expect, it } from 'vitest'
import { MenuCommand } from '@shared/types.js'
import { EditorCommand, asEditorCommand, runsWhileLocked } from './editor-commands.js'

describe('editor commands in read-only mode', () => {
  it('only search, counting, bookmarks, navigation, opening the equation, configuring assistants and the assistant panel run with the document locked', () => {
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
        EditorCommand.AiProviders,
        EditorCommand.AiAssistant,
      ].sort(),
    )
  })

  it('every table action is an edit', () => {
    for (const command of Object.values(EditorCommand).filter((value) => value.startsWith('table-'))) {
      expect(runsWhileLocked(command)).toBe(false)
    }
  })
})

describe('menu hand-off to the editor', () => {
  it('recognizes every editor command by name, and only those', () => {
    for (const command of Object.values(EditorCommand)) expect(asEditorCommand(command)).toBe(command)
    expect(asEditorCommand(MenuCommand.Save)).toBeNull()
  })
})
