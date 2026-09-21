import { describe, expect, it } from 'vitest'
import { lineFactorOf } from './line-metrics.js'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  FirstLineKind,
  INDENT_STEP_MM,
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

  it('a entrelinha simples é "normal" no bloco que nasceu no editor', () => {
    // O espaçamento simples do Word é a altura que a fonte pede, e nenhum fator
    // o imita. No bloco sem nada declarado a forma continua sendo `normal`: trocá-la
    // por um número reescreveria o bloco sem mudar uma linha do que se vê.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT).lineHeight).toBe('normal')
    expect(paragraphDraftFrom({ lineHeight: 'normal' }).lineSpacingKind).toBe(LineSpacingKind.Single)
  })

  it('a entrelinha em pontos volta como pontos', () => {
    const draft = paragraphDraftFrom({ lineHeight: '14pt' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.AtLeast)
    expect(draft.lineSpacingValue).toBe(14)
    expect(paragraphAttrsFrom(draft).lineHeight).toBe('14pt')
  })

  it('o fator do diálogo é o do Word, e o atributo é o do CSS', () => {
    // O atributo guarda o fator **já multiplicado** pela altura natural da fonte,
    // porque é isso que o CSS mede e é a forma que o leitor produz. Enquanto o
    // diálogo lia esse número como se fosse o do Word, "1,5" na tela gravava 1,23
    // linha no arquivo — 18 % de erro em Calibri.
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
    // A prova da ida e volta: o gravador divide pela mesma altura, e o número que
    // chega ao `w:line` é o que a pessoa escolheu.
    expect(lineFactorOf(Number(attrs.lineHeight), fontFamily)).toBe(factor)
    expect(paragraphDraftFrom({ fontFamily, lineHeight: attrs.lineHeight }).lineSpacingValue).toBe(factor)
  })

  it('o parágrafo importado sem entrelinha declarada abre como Simples', () => {
    // O leitor escreve o fator 1 já multiplicado — `1.2207` em Calibri —, e era
    // isso que o diálogo mostrava como "Múltiplo 1,2207": a opção "Simples" nunca
    // aparecia em documento nenhum.
    const draft = paragraphDraftFrom({ fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' })

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Single)
    // E aplicar de novo devolve o mesmo atributo: bloco intocado não é bloco
    // reescrito.
    expect(
      paragraphAttrsFrom(draft, { fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' }).lineHeight,
    ).toBe('1.2207')
  })

  it('escolher Simples sobre uma entrelinha declarada chega ao arquivo', () => {
    // Fonte que a tabela não conhece: para o fator 1 o leitor escreveria `normal`,
    // e `normal` é o único valor que o gravador **não** grava — deixaria o
    // `w:line` de 1,5 linha de pé. Silêncio sobre medida declarada é perda
    // silenciosa, então aqui sai número.
    const attrs = paragraphAttrsFrom(
      { ...DEFAULT_PARAGRAPH_DRAFT, lineSpacingKind: LineSpacingKind.Single },
      { fontFamily: 'Aptos', lineHeight: '1.7249' },
    )

    expect(attrs.lineHeight).toBe('1.1499')
    expect(lineFactorOf(Number(attrs.lineHeight), 'Aptos')).toBe(1)
  })

  it('o fator mínimo cabe na faixa do gravador', () => {
    // 0,51 era lido como medida de CSS e virava 0,42 de fator em Calibri — fora de
    // (0,5; 4), o que o gravador registra como perda e não escreve. Agora o número
    // do diálogo é o do Word, e o mínimo é mínimo de verdade.
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
    // A barra e os atalhos `Ctrl+1`, `Ctrl+5` e `Ctrl+2` usam a mesma conversão do
    // diálogo: sem ela a barra mostrava "1,8311" onde o Word mostra "1,5".
    const calibri = { fontFamily: 'Calibri, sans-serif', lineHeight: '1.8311' }

    expect(lineSpacingChoiceOf(calibri)).toBe('1.5')
    expect(lineSpacingChoiceOf({ fontFamily: 'Calibri, sans-serif', lineHeight: '1.2207' })).toBe('')
    expect(lineSpacingChoiceOf({ lineHeight: '14pt' })).toBe('14pt')

    expect(lineHeightAttrFrom('1.5', { fontFamily: 'Calibri, sans-serif' })).toBe('1.8311')
    expect(lineHeightAttrFrom('', calibri)).toBe('1.2207')
    expect(lineHeightAttrFrom('14pt', calibri)).toBe('14pt')
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

    expect(lineFactorOf(Number(attrs.lineHeight), null)).toBe(MAX_LINE_FACTOR)
  })

  it('o alinhamento só é escrito quando alguém o escolheu', () => {
    // O leitor omite `textAlign` no parágrafo sem `w:jc`. Escrevê-lo sempre fazia
    // "Aplicar" acrescentar `w:jc left` e marcar como alterado um bloco em que
    // ninguém tocou — o mesmo cuidado que os recuos já tinham.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, {}).textAlign).toBeNull()

    // Mas num parágrafo que **declara** alinhamento, escolher "À esquerda" é uma
    // decisão: apagar o atributo deixaria o estilo justificar de volta.
    expect(paragraphAttrsFrom(DEFAULT_PARAGRAPH_DRAFT, { textAlign: 'justify' }).textAlign).toBe(
      TextAlignment.Left,
    )
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
