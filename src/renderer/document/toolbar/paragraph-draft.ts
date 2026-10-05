import { LineSpacingKind, type ParagraphDraft } from '@services/document/paragraph-format.js'

/** Field groups receive this function: "Apply" needs to see the whole draft. */
export type DraftChange = <K extends keyof ParagraphDraft>(key: K, value: ParagraphDraft[K]) => void

/** The three common factors are ready-made options; "Multiple" is for the rest. */
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
