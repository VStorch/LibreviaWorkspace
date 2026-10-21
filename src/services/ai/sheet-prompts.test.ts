import { describe, expect, it } from 'vitest'
import { Language } from '@shared/i18n/language.js'
import { AiSheetTask } from '@shared/types.js'
import { buildSheetPrompt } from './sheet-prompts.js'

const CONTEXT = ['Sheet: Vendas', 'Active cell: B4', 'Cells:', 'B3: 2', 'B2: 3', 'A2: Caneta'].join('\n')
const BASE = { language: Language.Portuguese, instruction: 'some os preços', context: CONTEXT }

describe('buildSheetPrompt', () => {
  it('to create a formula, lists the functions in the stored dialect and asks for the formula only', () => {
    const prompt = buildSheetPrompt({ ...BASE, task: AiSheetTask.Formula }, 10_000)!
    expect(prompt.system).toContain('SOMA')
    expect(prompt.system).toContain('semicolon')
    expect(prompt.system).toContain('formula only')
    expect(prompt.messages[0]?.text).toBe(`Request: some os preços\n\n${CONTEXT}`)
  })

  it('to explain, sends the formula and asks for Markdown in the person language', () => {
    const request = { ...BASE, task: AiSheetTask.Explain, instruction: '', formula: '=SOMA(B2:B3)' }
    const prompt = buildSheetPrompt(request, 10_000)!
    expect(prompt.system).toContain('Brazilian Portuguese')
    expect(prompt.system).toContain('Markdown')
    expect(prompt.messages[0]?.text).toMatch(/^Formula: =SOMA\(B2:B3\)\n\nSheet: Vendas/)
  })

  it('drops the farthest cells to fit the provider context, discounting the instructions', () => {
    const request = { ...BASE, task: AiSheetTask.Formula }
    const system = buildSheetPrompt(request, Number.MAX_SAFE_INTEGER)!.system.length
    const prompt = buildSheetPrompt(request, system + 70)!
    const text = prompt.messages[0]!.text
    expect(text).toContain('B3: 2')
    expect(text).not.toContain('A2: Caneta')
  })

  it('to fill, asks for a JSON array of the right size and never cuts rows', () => {
    const request = { ...BASE, task: AiSheetTask.Fill, rowCount: 3 }
    expect(buildSheetPrompt(request, 10_000)!.system).toContain('JSON array only, with exactly 3 items')
    expect(buildSheetPrompt(request, 70)).toBeNull()
    const withoutInstructions = CONTEXT.length + BASE.instruction.length
    expect(buildSheetPrompt(request, withoutInstructions)).toBeNull()
  })
})
