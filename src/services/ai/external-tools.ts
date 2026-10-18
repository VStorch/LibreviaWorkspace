import { APPLIED_DOCUMENT_TOOLS } from './document-tools.js'
import { SHEET_TOOLS } from './sheet-tools.js'

/** What a client outside the app gets: the open file is a text document or a spreadsheet. */
export const EXTERNAL_TOOLS = [...APPLIED_DOCUMENT_TOOLS, ...SHEET_TOOLS]

export const EXTERNAL_INSTRUCTIONS =
  'Librevia has one file open at a time. The paragraph tools work on an open text document; the cell tools (list_sheets, read_cells, find_cells, write_cells) on an open spreadsheet. A tool for the other kind says which file is open.'
