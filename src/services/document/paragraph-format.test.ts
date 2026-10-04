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

describe('formatação de parágrafo', () => {
  it('o bloco sem nada declarado abre o diálogo no padrão', () => {
    // Bloco novo: todo atributo nulo. O diálogo não abre com `NaN`.
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

  it('o deslocamento é a mesma medida com o sinal trocado', () => {
    // O recuo pendente é `text-indent` negativo no CSS e duas perguntas no diálogo.
    const draft = paragraphDraftFrom({ firstLineMm: -6.4 })
    expect(draft.firstLineKind).toBe(FirstLineKind.Hanging)
    expect(draft.firstLineMm).toBe(6.4)

    expect(paragraphAttrsFrom(draft).firstLineMm).toBe(-6.4)
  })

  it('primeira linha e deslocamento não saem juntos', () => {
    const attrs = paragraphAttrsFrom({
      ...DEFAULT_PARAGRAPH_DRAFT,
      firstLineKind: FirstLineKind.Indent,
      firstLineMm: 12.5,
    })

    expect(attrs.firstLineMm).toBe(12.5)
  })

  it('recuo zerado sai como ausência, e não como zero', () => {
    // Nulo e ausente são o mesmo para a impressão digital; `0` explícito reescreveria o bloco.
    const attrs = paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT)

    expect(attrs.indentMm).toBeNull()
    expect(attrs.indentRightMm).toBeNull()
    expect(attrs.firstLineMm).toBeNull()
    expect(attrs.keepNext).toBeNull()
  })

  it('o passo de Ctrl+] entra no campo como milímetro', () => {
    // As duas origens de recuo somam no gravador, e o campo mostra a soma.
    const draft = paragraphDraftFrom({ indentMm: 5, indent: 2 })

    expect(draft.indentLeftMm).toBe(Math.round((5 + 2 * INDENT_STEP_MM) * 10) / 10)
    // E volta como medida única, com o nível zerado: senão o recuo dobraria.
    expect(paragraphAttrsFrom(draft).indent).toBe(0)
  })

  it('a entrelinha simples do bloco que nasceu no editor continua calada', () => {
    // "Simples" no bloco sem nada declarado deixa o atributo ausente: um número
    // reescreveria o bloco sem mudar o que se vê.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT).lineHeight).toBeNull()
    // Vindo de outra entrelinha é medida explícita: `normal` o gravador não grava.
    expect(
      paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, { lineHeight: '1.5' }, { lineHeight: '1.5' }).lineHeight,
    ).toBe('1.1499')
    expect(paragraphDraftFrom({ lineHeight: 'normal' }).lineSpacingKind).toBe(LineSpacingKind.Single)
  })

  it('a entrelinha em pontos volta como pontos', () => {
    const draft = paragraphDraftFrom({ lineHeight: '14pt' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.AtLeast)
    expect(draft.lineSpacingValue).toBe(14)
    expect(paragraphAttrsFrom(draft).lineHeight).toBe('14pt')
  })

  it('o fator do diálogo é o do Word, e o atributo é o do CSS', () => {
    // O atributo guarda o fator já multiplicado pela altura natural da fonte, como o
    // CSS mede e o leitor produz; o diálogo mostra o do Word.
    const calibri = { fontFamily: 'Calibri, sans-serif' }
    const draft = paragraphDraftFrom({ ...calibri, lineHeight: '1.8311' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Multiple)
    expect(draft.lineSpacingValue).toBe(1.5)
    // E de volta pelo mesmo caminho, que é o que o gravador vai dividir.
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
    // Um caso por fator e por altura natural: é a conta que o gravador desfaz
    // (`ParagraphFormat.ApplyLineHeight`), e o que o Word acaba lendo.
    const attrs = paragraphAttrsFrom(
      { ...DEFAULT_PARAGRAPH_DRAFT, lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: factor },
      { fontFamily },
    )

    expect(attrs.lineHeight).toBe(css)
    // O gravador divide pela mesma altura, e o `w:line` recebe o que a pessoa escolheu.
    expect(lineFactorOf(Number(attrs.lineHeight), fontFamily)).toBe(factor)
    expect(paragraphDraftFrom({ fontFamily, lineHeight: attrs.lineHeight }).lineSpacingValue).toBe(factor)
  })

  it('o parágrafo importado sem entrelinha declarada abre como Simples', () => {
    // O fator 1 já multiplicado (`1.2207` em Calibri) aparece como "Simples".
    const draft = paragraphDraftFrom({ fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Single)
    // E aplicar de novo devolve o mesmo atributo.
    expect(
      paragraphAttrsFrom(draft, { fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' }).lineHeight,
    ).toBe('1.2207')
  })

  it('escolher Simples sobre uma entrelinha declarada chega ao arquivo', () => {
    // Fonte desconhecida: `normal` deixaria de pé o `w:line` antigo, então sai número.
    const attrs = paragraphAttrsFrom(
      { ...DEFAULT_PARAGRAPH_DRAFT, lineSpacingKind: LineSpacingKind.Single },
      { fontFamily: 'Aptos', lineHeight: '1.7249' },
    )

    expect(attrs.lineHeight).toBe('1.1499')
    expect(lineFactorOf(Number(attrs.lineHeight), 'Aptos')).toBe(1)
  })

  it('o fator mínimo cabe na faixa do gravador', () => {
    // O número do diálogo é o do Word: 0,51 como medida de CSS seria 0,42 de fator,
    // fora de (0,5; 4).
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

  it('o seletor rápido da barra fala em linha, e não em CSS', () => {
    // A barra e `Ctrl+1`, `Ctrl+5` e `Ctrl+2` usam a conversão do diálogo.
    const calibri = { fontFamily: 'Calibri, sans-serif', lineHeight: '1.8311' }

    expect(lineSpacingChoiceOf(calibri)).toBe('1.5')
    expect(lineSpacingChoiceOf({ fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' })).toBe('')
    expect(lineSpacingChoiceOf({ lineHeight: '14pt' })).toBe('14pt')

    expect(lineHeightAttrFrom('1.5', { fontFamily: 'Calibri, sans-serif' })).toBe('1.8311')
    expect(lineHeightAttrFrom('', calibri)).toBe('1.2207')
    expect(lineHeightAttrFrom('14pt', calibri)).toBe('14pt')
  })

  it('o fator fica na faixa que o gravador aceita', () => {
    // Fora de (0,5; 4) o gravador registra perda: o diálogo nem oferece.
    const attrs = paragraphAttrsFrom({
      ...DEFAULT_PARAGRAPH_DRAFT,
      lineSpacingKind: LineSpacingKind.Multiple,
      lineSpacingValue: 12,
    })

    expect(lineFactorOf(Number(attrs.lineHeight), null)).toBe(MAX_LINE_FACTOR)
  })

  it('o alinhamento só é escrito quando alguém o escolheu', () => {
    // O leitor omite `textAlign` sem `w:jc`, e "Aplicar" não o acrescenta.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, {}).textAlign).toBeNull()

    // Num parágrafo que declara alinhamento, "À esquerda" é decisão: o estilo justificaria de volta.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, { textAlign: 'justify' }).textAlign).toBe(
      TextAlignment.Left,
    )
  })

  it('alinhamento que o modelo não conhece vira esquerda', () => {
    expect(paragraphDraftFrom({ textAlign: 'start' }).align).toBe(TextAlignment.Left)
    expect(paragraphDraftFrom({ textAlign: 'justify' }).align).toBe(TextAlignment.Justify)
  })

  it('campo vazio ou fora de faixa reprova o formulário', () => {
    // `NaN` do campo apagado gravaria `w:before="NaN"`, que o Word recusa.
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

  describe('contra o estilo', () => {
    // O bloco carrega só a formatação direta; o resto é do estilo, e o diálogo
    // mostra o que se vê — `effectiveAttrs` —, mas grava só o que mudou.
    const attrs = { styleId: 'Citacao', spaceAfter: 6, indent: 0 }
    const effective = {
      ...attrs,
      spaceBefore: 0,
      indentMm: 12.7,
      lineHeight: '1.3174',
      fontFamily: 'Calibri, sans-serif',
      keepNext: true,
    }

    it('OK sem mudança devolve cada atributo como estava', () => {
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

    it('grava só o campo que mudou, e o zero que desfaz o estilo é explícito', () => {
      const draft = { ...paragraphDraftFrom(effective), indentLeftMm: 0, keepNext: false }
      const result = paragraphAttrsFrom(draft, attrs, effective)

      expect(result.indentMm).toBe(0)
      expect(result.keepNext).toBe(false)
      expect(result.lineHeight).toBeNull()
      expect(result.spaceAfter).toBe(6)
    })

    it('a entrelinha nova é medida na fonte que o estilo dá ao bloco', () => {
      const draft = { ...paragraphDraftFrom(effective), lineSpacingValue: 1.5 }

      expect(paragraphAttrsFrom(draft, attrs, effective).lineHeight).toBe('1.8311')
    })
  })
})
