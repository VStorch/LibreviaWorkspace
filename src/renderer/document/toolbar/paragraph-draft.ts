import { LineSpacingKind, type ParagraphDraft } from '@services/document/paragraph-format.js'

/** Os grupos de campos recebem esta função: "Aplicar" precisa ver o rascunho inteiro. */
export type DraftChange = <K extends keyof ParagraphDraft>(key: K, value: ParagraphDraft[K]) => void

/** Os três fatores comuns são opções prontas; "Múltiplo" é para o resto. */
export function lineSpacingChoice(draft: ParagraphDraft): string {
  if (draft.lineSpacingKind === LineSpacingKind.Single) return 'single'
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast) return 'at-least'

  const exact = String(draft.lineSpacingValue)
  return ['1.15', '1.5', '2'].includes(exact) ? exact : 'multiple'
}

export function lineSpacingFrom(
  choice: string,
): Pick<ParagraphDraft, 'lineSpacingKind' | 'lineSpacingValue'> {
  if (choice === 'single') {
    return { lineSpacingKind: LineSpacingKind.Single, lineSpacingValue: 1.15 }
  }
  if (choice === 'at-least') {
    return { lineSpacingKind: LineSpacingKind.AtLeast, lineSpacingValue: 14 }
  }
  if (choice === 'multiple') {
    return { lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: 1.15 }
  }

  return { lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: Number(choice) }
}

export function isCustomLineSpacing(draft: ParagraphDraft): boolean {
  return draft.lineSpacingKind === LineSpacingKind.AtLeast || lineSpacingChoice(draft) === 'multiple'
}
