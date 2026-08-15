/**
 * Recálculo da pasta de trabalho.
 *
 * O problema não é calcular: é **em que ordem**. Se `A1` contém `=B1+1` e `B1`
 * contém `=C1*2`, calcular na ordem em que as células aparecem daria a A1 o
 * valor velho de B1 — e o erro apareceria só na segunda vez que alguém mexesse
 * na planilha, o que é o pior tipo de erro.
 *
 * A solução é a de sempre: montar o grafo de dependências e percorrê-lo em
 * ordem topológica. O que a ordem topológica não resolve é o ciclo, e é por isso
 * que a detecção dele vem junto, e não depois.
 */

import type { Cell, Sheet, WorkbookModel } from '../model.js'
import { walk, type Node } from './ast.js'
import { FormulaError } from './errors.js'
import { evaluate, type EvalContext } from './evaluate.js'
import { tryParseFormula } from './parse.js'
import type { CellRef } from './references.js'
import type { Scalar } from './values.js'

/** Uma célula com fórmula, já analisada. */
interface FormulaCell {
  readonly sheet: number
  readonly row: number
  readonly column: number
  /** A referência A1 como está no mapa, para gravar o resultado de volta. */
  readonly ref: string
  readonly node: Node | null
}

const key = (sheet: number, row: number, column: number): string => `${sheet}|${row}|${column}`

/**
 * Posição → chave numérica do índice de valores.
 *
 * Chave numérica, e não a referência `"A1"`: uma fórmula como `SOMA(A1:A10000)`
 * faz dez mil leituras, e montar dez mil textos a cada uma custava mais que a
 * conta inteira. O limite de colunas é o do Excel, então a chave nunca colide.
 */
const COLUMN_SPAN = 16_384
const at = (row: number, column: number): number => row * COLUMN_SPAN + column

/**
 * Recalcula todas as fórmulas e devolve a pasta com os valores atualizados.
 *
 * Planilhas sem fórmula voltam **como o mesmo objeto**: o React compara por
 * identidade, e recriá-las faria a grade inteira redesenhar a cada tecla.
 */
export function recalculate(workbook: WorkbookModel, now: () => Date = () => new Date()): WorkbookModel {
  const cells = collect(workbook)
  if (cells.length === 0) return workbook

  const byName = sheetsByName(workbook)
  const index = workbook.sheets.map(indexOf)
  const results: { cell: FormulaCell; value: Scalar }[] = []

  for (const cell of order(cells, byName)) {
    // Fórmula que não fecha só chega aqui num arquivo editado à mão: a
    // interface recusa antes de gravar.
    const value =
      cell.node === null
        ? FormulaError.Value
        : // O contexto é montado por célula porque uma referência sem nome de
          // planilha aponta para a planilha da **fórmula**, que muda a cada uma.
          evaluate(cell.node, contextFor(cell.sheet, index, byName, now))

    // Escrever no índice é o que faz a ordem topológica valer: quem vier depois
    // e depender desta célula já lê o valor novo.
    index[cell.sheet]?.set(at(cell.row, cell.column), value)
    results.push({ cell, value })
  }

  return apply(workbook, results)
}

/** Valores de uma planilha por chave numérica, para leitura rápida. */
function indexOf(sheet: Sheet): Map<number, Scalar> {
  const values = new Map<number, Scalar>()

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const position = positionOf(ref)
    if (position === null || cell.value === undefined) continue
    values.set(at(position.row, position.column), cell.value)
  }

  return values
}

function sheetsByName(workbook: WorkbookModel): ReadonlyMap<string, number> {
  const byName = new Map<string, number>()
  // O Excel não diferencia maiúsculas em nome de planilha, e o usuário digita
  // `plan1!A1` esperando que funcione.
  for (const [index, sheet] of workbook.sheets.entries()) byName.set(sheet.name.toUpperCase(), index)
  return byName
}

function collect(workbook: WorkbookModel): FormulaCell[] {
  const cells: FormulaCell[] = []

  for (const [sheet, model] of workbook.sheets.entries()) {
    for (const [ref, cell] of Object.entries(model.cells)) {
      if (cell.formula === undefined) continue

      const position = positionOf(ref)
      if (position === null) continue

      cells.push({ sheet, ...position, ref, node: tryParseFormula(cell.formula) })
    }
  }

  return cells
}

/**
 * O contexto de avaliação lê do índice, que já foi atualizado pelas fórmulas
 * calculadas antes desta.
 *
 * É o que faz a ordem topológica valer: quando `A1` pergunta por `B1`, ou `B1`
 * já foi calculada nesta passada e o índice tem o valor novo, ou ela não é
 * fórmula e o valor do arquivo é o definitivo.
 */
function contextFor(
  own: number,
  index: readonly Map<number, Scalar>[],
  byName: ReadonlyMap<string, number>,
  now: () => Date,
): EvalContext {
  return {
    now,
    valueAt: (ref: CellRef): Scalar => {
      const sheet = sheetIndexOf(ref, byName, own)
      // Nome de planilha que não existe: a aba foi excluída ou renomeada.
      if (sheet === null) return FormulaError.Ref

      return index[sheet]?.get(at(ref.row, ref.column)) ?? null
    },
  }
}

/**
 * A qual planilha uma referência pertence.
 *
 * Sem nome, é a planilha da própria fórmula. Como o recálculo percorre a pasta
 * inteira, a "própria" muda a cada célula — por isso o contexto é montado por
 * planilha, e não uma vez só.
 */
function sheetIndexOf(ref: CellRef, byName: ReadonlyMap<string, number>, fallback: number): number | null {
  if (ref.sheet === undefined) return fallback
  return byName.get(ref.sheet.toUpperCase()) ?? null
}

/**
 * Ordem de cálculo, com os ciclos já marcados.
 *
 * Percurso em profundidade com pilha explícita, e não recursão: uma coluna de
 * dez mil fórmulas encadeadas estouraria a pilha do JavaScript, e o usuário
 * veria o aplicativo morrer sem explicação.
 */
function order(cells: readonly FormulaCell[], byName: ReadonlyMap<string, number>): FormulaCell[] {
  const byKey = new Map<string, FormulaCell>()
  for (const cell of cells) byKey.set(key(cell.sheet, cell.row, cell.column), cell)

  const dependencies = new Map<string, string[]>()
  for (const cell of cells) {
    dependencies.set(key(cell.sheet, cell.row, cell.column), dependenciesOf(cell, byKey, byName))
  }

  const sorted: FormulaCell[] = []
  const state = new Map<string, 'visiting' | 'done'>()
  const circular = new Set<string>()

  for (const start of byKey.keys()) {
    if (state.has(start)) continue

    // Cada quadro guarda em que dependência parou, para retomar de onde saiu.
    const stack: { at: string; next: number }[] = [{ at: start, next: 0 }]
    state.set(start, 'visiting')

    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!
      const deps = dependencies.get(frame.at) ?? []

      if (frame.next >= deps.length) {
        stack.pop()
        state.set(frame.at, 'done')
        const cell = byKey.get(frame.at)
        if (cell !== undefined && !circular.has(frame.at)) sorted.push(cell)
        continue
      }

      const dependency = deps[frame.next++]!
      const seen = state.get(dependency)

      if (seen === 'done') continue
      if (seen === 'visiting') {
        // Fechou o ciclo: tudo que está na pilha a partir dele participa.
        const from = stack.findIndex((entry) => entry.at === dependency)
        for (const entry of stack.slice(from)) circular.add(entry.at)
        continue
      }

      state.set(dependency, 'visiting')
      stack.push({ at: dependency, next: 0 })
    }
  }

  // As circulares entram **primeiro**, já com o veredito. Nenhuma delas pode
  // ser calculada, e pôr o veredito antes é o que faz quem depende de uma
  // circular herdar o erro pela propagação normal, em vez de ler um valor velho.
  const broken: FormulaCell[] = []
  for (const at of circular) {
    const cell = byKey.get(at)
    if (cell !== undefined) broken.push({ ...cell, node: { kind: 'error', value: FormulaError.Circular } })
  }

  return [...broken, ...sorted]
}

/**
 * De quais **outras fórmulas** esta célula depende.
 *
 * Só as fórmulas importam para a ordem: depender de uma célula com número
 * digitado não impõe restrição nenhuma. Por isso os intervalos são cruzados
 * contra a lista de fórmulas em vez de percorridos célula a célula — senão
 * `SOMA(A1:A10000)` custaria dez mil passos para descobrir que depende de duas.
 */
function dependenciesOf(
  cell: FormulaCell,
  byKey: ReadonlyMap<string, FormulaCell>,
  byName: ReadonlyMap<string, number>,
): string[] {
  if (cell.node === null) return []

  const found = new Set<string>()

  for (const node of walk(cell.node)) {
    if (node.kind === 'reference') {
      const sheet = sheetIndexOf(node.ref, byName, cell.sheet)
      if (sheet === null) continue

      const at = key(sheet, node.ref.row, node.ref.column)
      if (byKey.has(at)) found.add(at)
      continue
    }

    if (node.kind !== 'range') continue

    const sheet = sheetIndexOf(node.from, byName, cell.sheet)
    if (sheet === null) continue

    for (const other of byKey.values()) {
      if (other.sheet !== sheet) continue
      if (other.row < node.from.row || other.row > node.to.row) continue
      if (other.column < node.from.column || other.column > node.to.column) continue
      found.add(key(other.sheet, other.row, other.column))
    }
  }

  // Uma fórmula que se referencia diretamente já é um ciclo.
  found.delete(key(cell.sheet, cell.row, cell.column))
  if (referencesItself(cell, byName)) found.add(key(cell.sheet, cell.row, cell.column))

  return [...found]
}

function referencesItself(cell: FormulaCell, byName: ReadonlyMap<string, number>): boolean {
  if (cell.node === null) return false

  for (const node of walk(cell.node)) {
    if (node.kind === 'reference') {
      const sheet = sheetIndexOf(node.ref, byName, cell.sheet)
      if (sheet === cell.sheet && node.ref.row === cell.row && node.ref.column === cell.column) return true
      continue
    }
    if (node.kind !== 'range') continue

    const sheet = sheetIndexOf(node.from, byName, cell.sheet)
    if (sheet !== cell.sheet) continue
    if (cell.row < node.from.row || cell.row > node.to.row) continue
    if (cell.column < node.from.column || cell.column > node.to.column) continue
    return true
  }

  return false
}

/**
 * Grava os valores calculados, preservando fórmula e formatação.
 *
 * Planilha cujo nenhum valor mudou volta como o **mesmo objeto**, e a pasta
 * também: o React compara por identidade, e recriá-las faria a grade inteira
 * redesenhar a cada tecla digitada em qualquer célula.
 */
function apply(
  workbook: WorkbookModel,
  results: readonly { cell: FormulaCell; value: Scalar }[],
): WorkbookModel {
  const changed = new Map<number, Record<string, Cell>>()

  for (const { cell, value } of results) {
    const sheet = workbook.sheets[cell.sheet]
    const previous = sheet?.cells[cell.ref]
    if (sheet === undefined || previous === undefined) continue

    // Fórmula nunca resulta em célula vazia: `=A1` sobre uma célula em branco
    // vale zero, como no Excel. Deixar vazio faria a fórmula desaparecer do
    // mapa esparso na próxima gravação, levando junto a própria fórmula.
    const calculated = value === null ? 0 : value
    if (calculated === previous.value) continue

    let cells = changed.get(cell.sheet)
    if (cells === undefined) {
      cells = { ...sheet.cells }
      changed.set(cell.sheet, cells)
    }
    cells[cell.ref] = { ...previous, value: calculated }
  }

  if (changed.size === 0) return workbook

  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet, index) => {
      const cells = changed.get(index)
      return cells === undefined ? sheet : { ...sheet, cells }
    }),
  }
}

function positionOf(ref: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([0-9]+)$/.exec(ref)
  if (match === null) return null

  let column = 0
  for (const letter of match[1]!) column = column * 26 + (letter.charCodeAt(0) - 64)

  return { row: Number.parseInt(match[2]!, 10) - 1, column: column - 1 }
}
