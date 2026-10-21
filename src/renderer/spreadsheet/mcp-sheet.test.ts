import { beforeEach, describe, expect, it } from 'vitest'
import { singleCell } from '@services/spreadsheet/edit.js'
import { createSheet, setCell } from '@services/spreadsheet/model.js'
import type { McpDocumentRequest } from '@shared/api.js'
import { MCP_DOCUMENT_REFUSED, McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'
import { answerSheetTool, useSheetMcpProposals } from './mcp-sheet.js'

const AUTHOR = 'Cliente (MCP)'
const SELECTION = singleCell(0, 0)

const request = (name: string, args: unknown): McpDocumentRequest => ({
  requestId: crypto.randomUUID(),
  ask: McpDocumentAsk.SheetTool,
  deadline: Date.now() + 60_000,
  author: AUTHOR,
  call: { name, arguments: JSON.stringify(args) },
})

const write = (): McpDocumentRequest => request('write_cells', { start: 'B1', rows: [['=A1*2']] })

describe('sheet tools requested by an MCP client', () => {
  beforeEach(() => {
    const sheet = setCell(createSheet('Planilha1'), 0, 0, { value: 3 })
    useWorkspace.setState({ workbook: { sheets: [sheet], activeSheet: 0 }, readOnly: false })
    useSheetMcpProposals.setState({ pending: [] })
  })

  it('the write stays pending with the author, without touching the sheet', () => {
    expect(answerSheetTool(write(), SELECTION)).toMatchObject({
      ask: McpDocumentAsk.SheetTool,
      isError: false,
    })
    expect(useSheetMcpProposals.getState().pending).toMatchObject([
      { sheet: 'Planilha1', row: 0, column: 1, author: AUTHOR },
    ])
    expect(useWorkspace.getState().workbook?.sheets[0]?.cells['B1']).toBeUndefined()
  })

  it('in read-only mode, reads but refuses the write without leaving a proposal', () => {
    useWorkspace.setState({ readOnly: true })
    expect(answerSheetTool(request('read_cells', { range: 'A1' }), SELECTION)).toMatchObject({
      isError: false,
      content: expect.stringContaining('A1: 3'),
    })
    expect(answerSheetTool(write(), SELECTION)).toMatchObject({
      ask: MCP_DOCUMENT_REFUSED,
      refusal: McpDocumentRefusal.ReadOnly,
    })
    expect(useSheetMcpProposals.getState().pending).toEqual([])
  })

  it('with no sheet on screen, refuses saying there is no sheet', () => {
    expect(answerSheetTool(write(), null)).toMatchObject({ refusal: McpDocumentRefusal.NoSpreadsheet })
  })
})
