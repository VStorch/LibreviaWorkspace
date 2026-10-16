import { EditCommand } from '@shared/types.js'

/** Through the edit menu's copy: the clipboard takes the formatting along with the text. */
export function copySelection(): void {
  void window.api.edit.run({ command: EditCommand.Copy })
}

export function copyContents(element: HTMLElement | null): void {
  if (element !== null) window.getSelection()?.selectAllChildren(element)
  copySelection()
}
