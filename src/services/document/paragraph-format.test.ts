import { describe, expect, it } from 'vitest'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  FirstLineKind,
  INDENT_STEP_MM,
  LineSpacingKind,
  MAX_LINE_FACTOR,
  TextAlignment,
  isValidParagraphDraft,
  paragraphAttrsFrom,
  paragraphDraftFrom,
} from './paragraph-format.js'

describe('formatação de parágrafo', () => {
  it('o bloco sem nada declarado abre o diálogo no padrão', () => {
    // Bloco novo do editor: todo atributo materializado como nulo. Sem isto o
    // diálogo abriria com `NaN` nos campos, e o primeiro "Aplicar" gravaria
    // medida inválida no documento.
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
    // No CSS o recuo pendente do Word é `text-indent` negativo — um número só.
    // No diálogo são duas perguntas, e a ida e a volta têm de fechar.
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
    // A impressão digital compara atributo nulo com atributo ausente como a
    // mesma coisa; um `0` explícito é diferença de forma, e faria a gravação
    // reescrever todo bloco que passasse pelo diálogo.
    const attrs = paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT)

    expect(attrs.indentMm).toBeNull()
    expect(attrs.indentRightMm).toBeNull()
    expect(attrs.firstLineMm).toBeNull()
    expect(attrs.keepNext).toBeNull()
  })

  it('o passo de Ctrl+] entra no campo como milímetro', () => {
    // As duas origens de recuo somam no gravador. Se o nível ficasse de fora, o
    // campo mostraria 0 mm para um parágrafo visivelmente recuado, e aplicar
    // qualquer outra coisa o desrecuaria.
    const draft = paragraphDraftFrom({ indentMm: 5, indent: 2 })

    expect(draft.indentLeftMm).toBe(Math.round((5 + 2 * INDENT_STEP_MM) * 10) / 10)
    // E volta como medida única, com o nível zerado: senão o recuo dobraria.
    expect(paragraphAttrsFrom(draft).indent).toBe(0)
  })

  it('a entrelinha simples é "normal", e não um fator', () => {
    // O espaçamento simples do Word é a altura que a fonte pede, e nenhum fator
    // o imita: 1,0 deixa cada linha meia altura mais apertada do que no Word.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT).lineHeight).toBe('normal')
    expect(paragraphDraftFrom({ lineHeight: 'normal' }).lineSpacingKind).toBe(LineSpacingKind.Single)
  })

  it('a entrelinha em pontos volta como pontos', () => {
    const draft = paragraphDraftFrom({ lineHeight: '14pt' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.AtLeast)
    expect(draft.lineSpacingValue).toBe(14)
    expect(paragraphAttrsFrom(draft).lineHeight).toBe('14pt')
  })

  it('a entrelinha em fator volta como fator', () => {
    const draft = paragraphDraftFrom({ lineHeight: '1.5' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Multiple)
    expect(paragraphAttrsFrom(draft).lineHeight).toBe('1.5')
  })

  it('o fator fica na faixa que o gravador aceita', () => {
    // Fora de (0,5; 4) o gravador registra perda em vez de escrever. Prender o
    // valor aqui é a diferença entre "o diálogo não oferece" e "o diálogo
    // oferece e o arquivo não recebe".
    const attrs = paragraphAttrsFrom({
      ...DEFAULT_PARAGRAPH_DRAFT,
      lineSpacingKind: LineSpacingKind.Multiple,
      lineSpacingValue: 12,
    })

    expect(attrs.lineHeight).toBe(String(MAX_LINE_FACTOR))
  })

  it('alinhamento que o modelo não conhece vira esquerda', () => {
    expect(paragraphDraftFrom({ textAlign: 'start' }).align).toBe(TextAlignment.Left)
    expect(paragraphDraftFrom({ textAlign: 'justify' }).align).toBe(TextAlignment.Justify)
  })

  it('campo vazio ou fora de faixa reprova o formulário', () => {
    // O `<input type="number">` devolve `NaN` quando a pessoa apaga o conteúdo,
    // e aplicar `NaN` gravaria `w:before="NaN"` — documento que o Word recusa.
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
})
