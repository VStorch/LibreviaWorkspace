import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import {
  allowExternalFile,
  fileFromArguments,
  forgetExternalFileRequest,
  isExternalFileRequested,
} from './external-files.js'

vi.mock('./window.js', () => ({ sendMenuCommand: () => {} }))

describe('fileFromArguments', () => {
  it.each(['relatorio.odt', 'contas.xlsx', 'manual.pdf', 'modelo.dotx'])('opens %s', (name) => {
    expect(fileFromArguments(['--flag', name], '/home/ana')).toBe(resolve('/home/ana', name))
  })

  it('ignores switches and what the app cannot open', () => {
    expect(fileFromArguments(['--inspect.docx', 'setup.exe', '.'], '/home/ana')).toBeUndefined()
  })
})

describe('allowExternalFile', () => {
  it('lets the renderer open a picked path once, without sending it a command', () => {
    allowExternalFile('/home/ana/a.docx')

    expect(isExternalFileRequested('/home/ana/a.docx')).toBe(true)
    forgetExternalFileRequest('/home/ana/a.docx')
    expect(isExternalFileRequested('/home/ana/a.docx')).toBe(false)
  })
})
