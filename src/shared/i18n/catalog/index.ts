import type { Entry } from '../message.js'
import { CHARS } from './chars.js'
import { COMMENTS } from './comments.js'
import { DIALOG } from './dialog.js'
import { DOCUMENT } from './document.js'
import { ERRORS } from './errors.js'
import { MENU } from './menu.js'
import { REFERENCES } from './references.js'
import { REVISIONS } from './revisions.js'
import { SHELL } from './shell.js'
import { SPREADSHEET } from './spreadsheet.js'
import { TABLE } from './table.js'
import { VIEW } from './view.js'

/**
 * The spread is shallow: a key repeated across areas vanishes silently, and the contract test
 * catches it.
 */
export const MESSAGES = {
  ...MENU,
  ...TABLE,
  ...VIEW,
  ...DOCUMENT,
  ...SPREADSHEET,
  ...SHELL,
  ...DIALOG,
  ...ERRORS,
  ...CHARS,
  ...REFERENCES,
  ...COMMENTS,
  ...REVISIONS,
} as const

export type MessageKey = keyof typeof MESSAGES

/** The contract test needs the areas apart. */
export const AREAS: Readonly<Record<string, Readonly<Record<string, Entry>>>> = {
  menu: MENU,
  table: TABLE,
  view: VIEW,
  document: DOCUMENT,
  spreadsheet: SPREADSHEET,
  shell: SHELL,
  dialog: DIALOG,
  errors: ERRORS,
  chars: CHARS,
  references: REFERENCES,
  comments: COMMENTS,
  revisions: REVISIONS,
}
