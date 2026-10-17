import { useState } from 'react'
import { describeRowsToFill, describeSheetAround } from '@services/ai/sheet-context.js'
import { singleCell, type Range } from '@services/spreadsheet/edit.js'
import { getCell, type Sheet, type WorkbookModel } from '@services/spreadsheet/model.js'
import type { MessageKey, Vars } from '@shared/i18n/index.js'
import { MAX_AI_FILL_ROWS, MAX_AI_INSTRUCTION_LENGTH, MAX_AI_SHEET_CONTEXT_LENGTH } from '@shared/limits.js'
import { AiSheetTask } from '@shared/types.js'
import { useT } from '../i18n.js'
import { AI_STATUS_LABELS } from '../ai/ai-status.js'
import { AiPanelFrame, AiProviderFields, useAiPanelChoice } from '../ai/AiPanelFrame.js'
import { SheetAiResult } from './SheetAiResult.js'
import { useSheetAi, type SheetAiAsk } from './useSheetAi.js'

const TASK_LABELS: Readonly<Record<AiSheetTask, MessageKey>> = {
  [AiSheetTask.Explain]: 'ai.sheet.explain',
  [AiSheetTask.Formula]: 'ai.sheet.formula',
  [AiSheetTask.Fill]: 'ai.sheet.fill',
}

const PLACEHOLDERS: Partial<Readonly<Record<AiSheetTask, MessageKey>>> = {
  [AiSheetTask.Formula]: 'ai.sheet.formulaPlaceholder',
  [AiSheetTask.Fill]: 'ai.sheet.fillPlaceholder',
}

interface Problem {
  readonly key: MessageKey
  readonly vars?: Vars
}

export interface SheetAiPanelProps {
  readonly workbook: WorkbookModel
  readonly range: Range
  readonly readOnly: boolean
  readonly configuring: boolean
  readonly onApply: (sheet: Sheet) => void
  readonly onClose: () => void
  readonly onConfigure: () => void
}

const rowsOf = (range: Range): number => range.toRow - range.fromRow + 1

function problemOf(task: AiSheetTask, sheet: Sheet, range: Range): Problem | null {
  if (task === AiSheetTask.Explain) {
    const formula = getCell(sheet, range.fromRow, range.fromColumn)?.formula
    return formula === undefined ? { key: 'ai.sheet.noFormula' } : null
  }
  if (task !== AiSheetTask.Fill) return null
  if (range.fromColumn !== range.toColumn) return { key: 'ai.sheet.oneColumn' }
  if (rowsOf(range) > MAX_AI_FILL_ROWS)
    return { key: 'ai.sheet.tooManyRows', vars: { max: MAX_AI_FILL_ROWS } }
  return null
}

type Question = Omit<SheetAiAsk, 'providerId' | 'model' | 'sheetIndex'>

function questionFor(task: AiSheetTask, sheet: Sheet, range: Range, instruction: string): Question {
  const active = { row: range.fromRow, column: range.fromColumn }
  if (task === AiSheetTask.Fill) {
    const context = describeRowsToFill(sheet, range)
    return { task, instruction, context, rowCount: rowsOf(range), target: range }
  }
  const context = describeSheetAround(sheet, active, MAX_AI_SHEET_CONTEXT_LENGTH)
  const target = singleCell(active.row, active.column)
  if (task === AiSheetTask.Formula) return { task, instruction, context, target }
  const formula = getCell(sheet, active.row, active.column)?.formula
  return { task, instruction: '', formula, context, target }
}

export function SheetAiPanel(props: SheetAiPanelProps): React.JSX.Element {
  const t = useT()
  const { workbook, range, readOnly } = props
  const choice = useAiPanelChoice(props.configuring)
  const ai = useSheetAi()
  const [task, setTask] = useState<AiSheetTask>(AiSheetTask.Formula)
  const [instruction, setInstruction] = useState('')
  const [tooLargeFor, setTooLargeFor] = useState<Range | null>(null)
  const sheet = workbook.sheets[workbook.activeSheet]!
  const tooLarge = task === AiSheetTask.Fill && tooLargeFor === range
  const problem = tooLarge ? { key: 'ai.sheet.tooLarge' as const } : problemOf(task, sheet, range)
  const needsInstruction = task !== AiSheetTask.Explain && instruction.trim() === ''
  const running = ai.run?.phase === 'running'

  const generate = (): void => {
    const { provider } = choice
    if (provider === undefined || problem !== null) return
    const question = questionFor(task, sheet, range, instruction)
    if (question.context.length > MAX_AI_SHEET_CONTEXT_LENGTH) return setTooLargeFor(range)
    const target = { providerId: provider.id, model: provider.defaultModel ?? '' }
    void ai.ask({ ...target, ...question, sheetIndex: workbook.activeSheet })
  }
  const status = ai.run === null ? undefined : AI_STATUS_LABELS[ai.run.phase]

  return (
    <AiPanelFrame choice={choice} onClose={props.onClose} onConfigure={props.onConfigure}>
      {readOnly && <p className="ai-pane__note">{t('ai.sheet.readOnly')}</p>}
      <AiProviderFields choice={choice} running={running} />
      <TaskFields task={task} onTask={setTask} instruction={instruction} onInstruction={setInstruction} />
      {problem !== null && <p className="ai-pane__note">{t(problem.key, problem.vars)}</p>}
      <div className="ai-pane__actions">
        {running ? (
          <button type="button" className="btn" onClick={ai.cancel}>
            {t('ai.panel.stop')}
          </button>
        ) : (
          <button
            type="button"
            className="btn btn--primary"
            disabled={choice.provider === undefined || problem !== null || needsInstruction}
            onClick={generate}
          >
            {t('ai.panel.generate')}
          </button>
        )}
      </div>
      {ai.run !== null && (
        <SheetAiResult run={ai.run} workbook={workbook} readOnly={readOnly} onApply={props.onApply} />
      )}
      {status !== undefined && (
        <p className="ai-pane__note" role="status">
          {t(status)}
        </p>
      )}
      {ai.run?.costUsd !== undefined && (
        <p className="ai-pane__note">{t('ai.panel.cost', { cost: ai.run.costUsd.toFixed(4) })}</p>
      )}
    </AiPanelFrame>
  )
}

interface TaskFieldsProps {
  readonly task: AiSheetTask
  readonly onTask: (task: AiSheetTask) => void
  readonly instruction: string
  readonly onInstruction: (instruction: string) => void
}

function TaskFields({ task, onTask, instruction, onInstruction }: TaskFieldsProps): React.JSX.Element {
  const t = useT()
  const placeholder = PLACEHOLDERS[task]
  return (
    <>
      <label className="popover__field">
        <span>{t('ai.panel.action')}</span>
        <select value={task} onChange={(event) => onTask(event.target.value as AiSheetTask)}>
          {Object.values(AiSheetTask).map((candidate) => (
            <option key={candidate} value={candidate}>
              {t(TASK_LABELS[candidate])}
            </option>
          ))}
        </select>
      </label>
      {placeholder !== undefined && (
        <label className="popover__field">
          <span>{t('ai.panel.instruction')}</span>
          <textarea
            rows={3}
            value={instruction}
            maxLength={MAX_AI_INSTRUCTION_LENGTH}
            placeholder={t(placeholder)}
            onChange={(event) => onInstruction(event.target.value)}
          />
        </label>
      )}
    </>
  )
}
