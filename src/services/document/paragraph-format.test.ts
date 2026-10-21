import { describe, expect, it } from 'vitest'
import { INDENT_STEP_MM } from '@services/units.js'
import { lineFactorOf } from './line-metrics.js'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  FirstLineKind,
  LineSpacingKind,
  MAX_LINE_FACTOR,
  MIN_LINE_FACTOR,
  TextAlignment,
  isValidParagraphDraft,
  lineHeightAttrFrom,
  lineSpacingChoiceOf,
  paragraphAttrsFrom,
  paragraphDraftFrom,
} from './paragraph-format.js'

describe('paragraph formatting', () => {
  it('a block with nothing declared opens the dialog at the default', () => {
    // A new block: every attribute null. The dialog does not open with `NaN`.
    const draft = paragraphDraftFrom({
      textAlign: null,
      spaceBefore: null,
      spaceAfter: null,
      lineHeight: null,
      indentMm: null,
      firstLineMm: null,
      keepNext: null,
    })

    expect(draft).toEqual(DEFAULT_PARAGRAPH_DRAFT)
  })

  it('hanging is the same measure with the sign flipped', () => {
    // Hanging indent is a negative CSS `text-indent` and two questions in the dialog.
    const draft = paragraphDraftFrom({ firstLineMm: -6.4 })
    expect(draft.firstLineKind).toBe(FirstLineKind.Hanging)
    expect(draft.firstLineMm).toBe(6.4)

    expect(paragraphAttrsFrom(draft).firstLineMm).toBe(-6.4)
  })

  it('first line and hanging do not come out together', () => {
    const attrs = paragraphAttrsFrom({
      ...DEFAULT_PARAGRAPH_DRAFT,
      firstLineKind: FirstLineKind.Indent,
      firstLineMm: 12.5,
    })

    expect(attrs.firstLineMm).toBe(12.5)
  })

  it('a zeroed indent comes out as absence, not as zero', () => {
    // Null and absent are the same to the fingerprint; an explicit `0` would rewrite the block.
    const attrs = paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT)

    expect(attrs.indentMm).toBeNull()
    expect(attrs.indentRightMm).toBeNull()
    expect(attrs.firstLineMm).toBeNull()
    expect(attrs.keepNext).toBeNull()
  })

  it('the Ctrl+] step enters the field as millimeters', () => {
    // Both indent sources add up in the writer, and the field shows the sum.
    const draft = paragraphDraftFrom({ indentMm: 5, indent: 2 })

    expect(draft.indentLeftMm).toBe(Math.round((5 + 2 * INDENT_STEP_MM) * 10) / 10)
    // And it comes back as a single measure, with the level reset: otherwise the indent would
    // double.
    expect(paragraphAttrsFrom(draft).indent).toBe(0)
  })

  it('single line spacing on a block born in the editor stays silent', () => {
    // "Single" on a block with nothing declared leaves the attribute absent: a number would rewrite
    // the block without changing what is visible.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT).lineHeight).toBeNull()
    // Coming from another spacing it is an explicit measure: the writer does not write `normal`.
    expect(
      paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, { lineHeight: '1.5' }, { lineHeight: '1.5' }).lineHeight,
    ).toBe('1.1499')
    expect(paragraphDraftFrom({ lineHeight: 'normal' }).lineSpacingKind).toBe(LineSpacingKind.Single)
  })

  it('line spacing in points comes back as points', () => {
    const draft = paragraphDraftFrom({ lineHeight: '14pt' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.AtLeast)
    expect(draft.lineSpacingValue).toBe(14)
    expect(paragraphAttrsFrom(draft).lineHeight).toBe('14pt')
  })

  it("the dialog factor is Word's, and the attribute is CSS's", () => {
    // The attribute stores the factor already multiplied by the font's natural height, as CSS
    // measures and the reader produces; the dialog shows Word's.
    const calibri = { fontFamily: 'Calibri, sans-serif' }
    const draft = paragraphDraftFrom({ ...calibri, lineHeight: '1.8311' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Multiple)
    expect(draft.lineSpacingValue).toBe(1.5)
    // And back the same way, which is what the writer will divide.
    expect(paragraphAttrsFrom(draft, calibri).lineHeight).toBe('1.8311')
  })

  it.each([
    ['Calibri, sans-serif', 1.15, '1.4038'],
    ['Calibri, sans-serif', 1.5, '1.8311'],
    ['Calibri, sans-serif', 2, '2.4414'],
    ['Times New Roman, serif', 1.5, '1.7249'],
    ['Courier New, monospace', 2, '2.2656'],
    ['Aptos', 1.5, '1.7249'],
  ])('em %s, %s linha grava %s de CSS', (fontFamily, factor, css) => {
    // One case per factor and natural height: the math the writer undoes
    // (`ParagraphFormat.ApplyLineHeight`), and what Word ends up reading.
    const attrs = paragraphAttrsFrom(
      { ...DEFAULT_PARAGRAPH_DRAFT, lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: factor },
      { fontFamily },
    )

    expect(attrs.lineHeight).toBe(css)
    // The writer divides by the same height, and `w:line` gets what the user chose.
    expect(lineFactorOf(Number(attrs.lineHeight), fontFamily)).toBe(factor)
    expect(paragraphDraftFrom({ fontFamily, lineHeight: attrs.lineHeight }).lineSpacingValue).toBe(factor)
  })

  it('an imported paragraph with no declared line spacing opens as Single', () => {
    // Factor 1 already multiplied (`1.2207` in Calibri) shows as "Single".
    const draft = paragraphDraftFrom({ fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Single)
    // And applying again gives the same attribute.
    expect(
      paragraphAttrsFrom(draft, { fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' }).lineHeight,
    ).toBe('1.2207')
  })

  it('choosing Single over a declared line spacing reaches the file', () => {
    // Unknown font: `normal` would leave the old `w:line` standing, so a number goes out.
    const attrs = paragraphAttrsFrom(
      { ...DEFAULT_PARAGRAPH_DRAFT, lineSpacingKind: LineSpacingKind.Single },
      { fontFamily: 'Aptos', lineHeight: '1.7249' },
    )

    expect(attrs.lineHeight).toBe('1.1499')
    expect(lineFactorOf(Number(attrs.lineHeight), 'Aptos')).toBe(1)
  })

  it("the minimum factor fits the writer's range", () => {
    // The dialog number is Word's: 0.51 as a CSS measure would be a 0.42 factor, outside (0.5; 4).
    const attrs = paragraphAttrsFrom(
      {
        ...DEFAULT_PARAGRAPH_DRAFT,
        lineSpacingKind: LineSpacingKind.Multiple,
        lineSpacingValue: MIN_LINE_FACTOR,
      },
      { fontFamily: 'Calibri, sans-serif' },
    )

    expect(lineFactorOf(Number(attrs.lineHeight), 'Calibri, sans-serif')).toBe(MIN_LINE_FACTOR)
  })

  it('the toolbar quick selector speaks in lines, not CSS', () => {
    // The toolbar and `Ctrl+1`, `Ctrl+5` and `Ctrl+2` use the dialog conversion.
    const calibri = { fontFamily: 'Calibri, sans-serif', lineHeight: '1.8311' }

    expect(lineSpacingChoiceOf(calibri)).toBe('1.5')
    expect(lineSpacingChoiceOf({ fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' })).toBe('')
    expect(lineSpacingChoiceOf({ lineHeight: '14pt' })).toBe('14pt')

    expect(lineHeightAttrFrom('1.5', { fontFamily: 'Calibri, sans-serif' })).toBe('1.8311')
    expect(lineHeightAttrFrom('', calibri)).toBe('1.2207')
    expect(lineHeightAttrFrom('14pt', calibri)).toBe('14pt')
  })

  it('the factor stays within the range the writer accepts', () => {
    // Outside (0.5; 4) the writer records a loss: the dialog does not even offer it.
    const attrs = paragraphAttrsFrom({
      ...DEFAULT_PARAGRAPH_DRAFT,
      lineSpacingKind: LineSpacingKind.Multiple,
      lineSpacingValue: 12,
    })

    expect(lineFactorOf(Number(attrs.lineHeight), null)).toBe(MAX_LINE_FACTOR)
  })

  it('alignment is written only when someone chose it', () => {
    // The reader omits `textAlign` without `w:jc`, and "Apply" does not add it.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, {}).textAlign).toBeNull()

    // In a paragraph that declares alignment, "Left" is a decision: the style would justify it
    // back.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, { textAlign: 'justify' }).textAlign).toBe(
      TextAlignment.Left,
    )
  })

  it('an alignment the model does not know becomes left', () => {
    expect(paragraphDraftFrom({ textAlign: 'start' }).align).toBe(TextAlignment.Left)
    expect(paragraphDraftFrom({ textAlign: 'justify' }).align).toBe(TextAlignment.Justify)
  })

  it('an empty or out-of-range field fails the form', () => {
    // A cleared field's `NaN` would write `w:before="NaN"`, which Word refuses.
    expect(isValidParagraphDraft(DEFAULT_PARAGRAPH_DRAFT)).toBe(true)
    expect(isValidParagraphDraft({ ...DEFAULT_PARAGRAPH_DRAFT, spaceBefore: Number.NaN })).toBe(false)
    expect(isValidParagraphDraft({ ...DEFAULT_PARAGRAPH_DRAFT, indentLeftMm: -3 })).toBe(false)
    expect(
      isValidParagraphDraft({
        ...DEFAULT_PARAGRAPH_DRAFT,
        lineSpacingKind: LineSpacingKind.AtLeast,
        lineSpacingValue: 0,
      }),
    ).toBe(false)
  })

  describe('against the style', () => {
    // The block carries only direct formatting; the rest is the style's, and the dialog shows what
    // is visible (`effectiveAttrs`) but saves only what changed.
    const attrs = { styleId: 'Citacao', spaceAfter: 6, indent: 0 }
    const effective = {
      ...attrs,
      spaceBefore: 0,
      indentMm: 12.7,
      lineHeight: '1.3174',
      fontFamily: 'Calibri, sans-serif',
      keepNext: true,
    }

    it('OK with no change returns each attribute as it was', () => {
      const result = paragraphAttrsFrom(paragraphDraftFrom(effective), attrs, effective)

      expect(result).toEqual({
        textAlign: null,
        spaceBefore: null,
        spaceAfter: 6,
        lineHeight: null,
        indentMm: null,
        indentRightMm: null,
        firstLineMm: null,
        keepNext: null,
        keepLines: null,
        widowControl: null,
        indent: 0,
      })
    })

    it('saves only the field that changed, and the zero that undoes the style is explicit', () => {
      const draft = { ...paragraphDraftFrom(effective), indentLeftMm: 0, keepNext: false }
      const result = paragraphAttrsFrom(draft, attrs, effective)

      expect(result.indentMm).toBe(0)
      expect(result.keepNext).toBe(false)
      expect(result.lineHeight).toBeNull()
      expect(result.spaceAfter).toBe(6)
    })

    it('the new line spacing is measured in the font the style gives the block', () => {
      const draft = { ...paragraphDraftFrom(effective), lineSpacingValue: 1.5 }

      expect(paragraphAttrsFrom(draft, attrs, effective).lineHeight).toBe('1.8311')
    })
  })
})
