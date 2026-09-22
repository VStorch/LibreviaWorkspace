import { describe, expect, it } from 'vitest'
import {
  CELL_BORDER_SIDES,
  CellBorderStyle,
  DEFAULT_TABLE_DRAFT,
  MAX_BORDER_PT,
  MAX_TABLE_COLUMNS,
  MAX_TABLE_ROWS,
  NO_CELL_BORDERS,
  cellBordersFromAttr,
  cellBordersToAttr,
  cellBordersToCss,
  isValidTableDraft,
  isValidTableSize,
  resolvedColumnWidths,
  tableDraftFrom,
  cellLookPatch,
  withBorderOnSides,
} from './table-format.js'

describe('borda e sombreamento de célula', () => {
  it('célula sem nada declarado abre o diálogo sem lado marcado', () => {
    // Célula recém-inserida: o ProseMirror materializa os dois atributos como
    // nulos. O diálogo mostra a verdade — nenhum lado declarado —, e o padrão
    // com os quatro lados é o que o botão "Restaurar padrão" oferece; sem isto o
    // diálogo diria que a célula tem bordas que ela não tem.
    expect(tableDraftFrom({ borders: null, shading: null })).toEqual({
      ...DEFAULT_TABLE_DRAFT,
      sides: { top: false, right: false, bottom: false, left: false },
    })
  })

  it('o texto do atributo tem uma escrita só', () => {
    // É o que a impressão digital compara: `0,50 pt` de um lado e `0.5 pt` do
    // outro descreveriam a mesma borda com dois textos, e toda tabela do
    // documento seria regenerada ao salvar.
    const borders = withBorderOnSides(NO_CELL_BORDERS, ['top', 'bottom'], {
      style: CellBorderStyle.Double,
      widthPt: 1.5,
      color: '#ff0000',
    })

    expect(cellBordersToAttr(borders)).toBe('top:double,1.5,#ff0000;bottom:double,1.5,#ff0000')
  })

  it('sem lado declarado o atributo é nulo, e não texto vazio', () => {
    // Nulo é o que o leitor omite e o schema materializa: ausente e nulo dizem a
    // mesma coisa, texto vazio diria "há borda, sem lado".
    expect(cellBordersToAttr(NO_CELL_BORDERS)).toBeNull()
  })

  it('a ida e a volta do atributo fecham', () => {
    const attr = 'top:single,0.5,#000000;left:dotted,2,#123456'
    expect(cellBordersToAttr(cellBordersFromAttr(attr))).toBe(attr)
  })

  it('texto estragado não derruba a célula', () => {
    // O atributo pode chegar de um HTML colado ou de um documento de outra
    // versão: o pior caso aceitável é a célula abrir sem borda nenhuma.
    expect(cellBordersFromAttr('lixo')).toEqual(NO_CELL_BORDERS)
    expect(cellBordersFromAttr(42)).toEqual(NO_CELL_BORDERS)
    expect(cellBordersFromAttr('top:enfeite,1,#000000')).toEqual(NO_CELL_BORDERS)
  })

  it('cor fora do formato do arquivo volta ao preto', () => {
    // `w:color` só aceita seis dígitos hexadecimais; `rgb(…)` faz o Word
    // declarar o documento danificado.
    expect(cellBordersFromAttr('top:single,1,rgb(0,0,0)').top?.color).toBe('#000000')
  })

  it('borda apagada de propósito vira largura zero no CSS', () => {
    // `w:val="nil"` **apaga** a borda que a tabela pediu. Com `border: none` e
    // sem largura, a borda da tabela reapareceria por baixo em `border-collapse`.
    const borders = withBorderOnSides(NO_CELL_BORDERS, ['bottom'], {
      style: CellBorderStyle.None,
      widthPt: 0.5,
      color: '#000000',
    })

    expect(cellBordersToCss(borders)).toBe('border-bottom:0')
  })

  it('o CSS sai em pontos, que é a unidade do arquivo', () => {
    const borders = withBorderOnSides(NO_CELL_BORDERS, ['right'], {
      style: CellBorderStyle.Dashed,
      widthPt: 2.25,
      color: '#aabbcc',
    })

    expect(cellBordersToCss(borders)).toBe('border-right:2.25pt dashed #aabbcc')
  })

  it('lado desmarcado no diálogo volta a nulo, e não a borda apagada', () => {
    // São escolhas diferentes: "não tenho opinião" deixa a borda da tabela
    // valer, "apagada" a remove. No arquivo, a ausência do elemento e `w:nil`.
    const cell = { borders: 'top:single,0.5,#000000;right:single,0.5,#000000', shading: null }
    const before = tableDraftFrom(cell)
    const after = { ...before, sides: { ...before.sides, right: false } }

    expect(cellLookPatch(cell, before, after)).toEqual({ borders: 'top:single,0.5,#000000', shading: null })
  })

  it('o sombreamento sai em minúscula, como o leitor o escreve', () => {
    const cell = { borders: null, shading: null }
    const before = tableDraftFrom(cell)
    const after = { ...before, shaded: true, shadingColor: '#D9D9D9' }
    expect(cellLookPatch(cell, before, after).shading).toBe('#d9d9d9')
  })

  it('o diálogo abre com os quatro lados que a célula tem', () => {
    const draft = tableDraftFrom({
      borders: 'top:double,3,#00ff00;right:double,3,#00ff00',
      shading: '#d9d9d9',
      columnWidthMm: 42,
      headerRow: true,
    })

    expect(draft.borderStyle).toBe(CellBorderStyle.Double)
    expect(draft.borderWidthPt).toBe(3)
    expect(draft.sides).toEqual({ top: true, right: true, bottom: false, left: false })
    expect(draft.shaded).toBe(true)
    expect(draft.columnWidthMm).toBe(42)
    expect(draft.headerRow).toBe(true)
  })

  it('quatro lados, nem mais nem menos', () => {
    // Um lado a mais aqui seria um lado que o `w:tcBorders` não tem.
    expect(CELL_BORDER_SIDES).toEqual(['top', 'right', 'bottom', 'left'])
  })
})

describe('largura das colunas como a tela as desenha', () => {
  it('tabela sem largura declarada divide a coluna de texto em partes iguais', () => {
    // É exatamente o que `table-layout: fixed` com `width: 100%` faz, e é o caso
    // da tabela recém-inserida: o diálogo tem de mostrar a medida que se vê.
    expect(resolvedColumnWidths([null, null, null], 600)).toEqual([200, 200, 200])
  })

  it('o que resta é dividido entre as colunas sem medida', () => {
    expect(resolvedColumnWidths([300, null, null], 600)).toEqual([300, 150, 150])
  })

  it('largura declarada em todas passa inteira', () => {
    // Sem tocar no total: a tabela mais larga que a coluna de texto é decisão de
    // quem arrastou a divisória, e o Word também a desenha estourando a margem.
    expect(resolvedColumnWidths([400, 500], 600)).toEqual([400, 500])
  })

  it('coluna nunca fica com medida que o gravador descarta', () => {
    // Zero e negativo fazem o gravador jogar a grade inteira fora — largura
    // parcial não é grade.
    expect(resolvedColumnWidths([null, null], 1).every((width) => width >= 1)).toBe(true)
  })
})

describe('validação do formulário de tabela', () => {
  it('espessura fora da faixa do arquivo não passa', () => {
    // `w:sz` mede em oitavos de ponto e cabe em um byte: acima de 31 pt o
    // documento sairia fora do esquema.
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderWidthPt: 0 })).toBe(false)
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderWidthPt: MAX_BORDER_PT + 1 })).toBe(false)
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderWidthPt: Number.NaN })).toBe(false)
  })

  it('largura de coluna não medida é válida — ela simplesmente não muda', () => {
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, columnWidthMm: null })).toBe(true)
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, columnWidthMm: 1 })).toBe(false)
  })

  it('cor que o arquivo não aceita reprova o formulário', () => {
    expect(isValidTableDraft({ ...DEFAULT_TABLE_DRAFT, borderColor: 'vermelho' })).toBe(false)
  })

  it('tamanho de tabela tem teto, e não é purismo', () => {
    // Cada célula é um parágrafo medido pela paginação: mil por mil pedidos por
    // engano num campo numérico travariam o editor.
    expect(isValidTableSize(3, 3)).toBe(true)
    expect(isValidTableSize(0, 3)).toBe(false)
    expect(isValidTableSize(2.5, 3)).toBe(false)
    expect(isValidTableSize(MAX_TABLE_ROWS + 1, 3)).toBe(false)
    expect(isValidTableSize(3, MAX_TABLE_COLUMNS + 1)).toBe(false)
  })
})

describe('o diálogo aplica só o que a pessoa mexeu', () => {
  // O Word grava `w:val="nil"` o tempo todo, e o rascunho do diálogo usa o
  // estilo da primeira borda declarada para os quatro lados. Aplicar o rascunho
  // inteiro apagava a borda de baixo de quem só trocou o sombreamento.
  const cell = { borders: 'top:none,0.5,#000000;bottom:single,1.5,#ff0000', shading: null }
  const before = tableDraftFrom(cell)

  it('trocar só o sombreamento não toca nas bordas', () => {
    const after = { ...before, shaded: true, shadingColor: '#ffff00' }
    expect(cellLookPatch(cell, before, after)).toEqual({
      borders: 'top:none,0.5,#000000;bottom:single,1.5,#ff0000',
      shading: '#ffff00',
    })
  })

  it('trocar a cor muda a cor de cada lado e mantém o estilo e a espessura de cada um', () => {
    const lados = { borders: 'top:single,0.5,#000000;bottom:double,3,#000000', shading: '#d9d9d9' }
    const antes = tableDraftFrom(lados)
    const after = { ...antes, borderColor: '#0000FF' }

    expect(cellLookPatch(lados, antes, after)).toEqual({
      borders: 'top:single,0.5,#0000ff;bottom:double,3,#0000ff',
      shading: '#d9d9d9',
    })
  })

  it('desmarcar um lado tira só ele, e marcar um lado novo o põe com o que o diálogo mostra', () => {
    const after = {
      ...before,
      sides: { ...before.sides, bottom: false, left: true },
    }

    expect(cellLookPatch(cell, before, after).borders).toBe('top:none,0.5,#000000;left:none,0.5,#000000')
  })

  it('nada mexido é nada aplicado', () => {
    expect(cellLookPatch(cell, before, before)).toEqual(cell)
  })

  it('em outra célula da seleção, o que mudou vale e o resto dela fica', () => {
    // O rascunho nasce da célula do cursor; a vizinha selecionada tem bordas
    // próprias, e só o campo alterado chega a ela.
    const vizinha = { borders: 'right:dotted,2,#00ff00', shading: '#eeeeee' }
    const after = { ...before, borderWidthPt: 1 }

    expect(cellLookPatch(vizinha, before, after)).toEqual({
      borders: 'right:dotted,1,#00ff00',
      shading: '#eeeeee',
    })
  })
})
