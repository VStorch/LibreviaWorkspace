import type { Entry } from '../message.js'
import { CHARS } from './chars.js'
import { DIALOG } from './dialog.js'
import { DOCUMENT } from './document.js'
import { ERRORS } from './errors.js'
import { MENU } from './menu.js'
import { SHELL } from './shell.js'
import { SPREADSHEET } from './spreadsheet.js'
import { TABLE } from './table.js'
import { VIEW } from './view.js'

/**
 * O catálogo inteiro, montado das áreas.
 *
 * Uma área por arquivo porque o catálogo cresce até virar a maior coisa do
 * repositório, e um arquivo de mil linhas é um arquivo em que ninguém acha nada
 * — e, pior, em que dois trabalhos simultâneos conflitam em toda linha. As
 * áreas seguem a interface: menu, aparência, documento, planilha, erros.
 *
 * O espalhamento é raso de propósito. Chave repetida entre duas áreas some em
 * silêncio, e é por isso que o teste de contrato do catálogo confere que nenhuma
 * se repete — o compilador não vê esse caso.
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
} as const

/**
 * Toda chave que existe.
 *
 * É isto que faz a varredura de tradução ser verificável em vez de confiável:
 * `t('menu.flie')` não compila, e uma chave removida do catálogo acusa cada
 * lugar que ainda a usava.
 */
export type MessageKey = keyof typeof MESSAGES

/** As áreas, por nome — o teste de contrato precisa delas separadas. */
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
}
