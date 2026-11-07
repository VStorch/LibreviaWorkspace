import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { fileFromArguments } from './external-files.js'

vi.mock('./window.js', () => ({ sendMenuCommand: () => {} }))

describe('fileFromArguments', () => {
  it.each(['relatorio.odt', 'contas.xlsx', 'manual.pdf', 'modelo.dotx'])('opens %s', (name) => {
    expect(fileFromArguments(['--flag', name], '/home/ana')).toBe(resolve('/home/ana', name))
  })

  it('ignores switches and what the app cannot open', () => {
    expect(fileFromArguments(['--inspect.docx', 'setup.exe', '.'], '/home/ana')).toBeUndefined()
  })
})
