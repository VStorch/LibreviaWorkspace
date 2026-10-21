import { describe, expect, it } from 'vitest'
import { adjustForColumns, adjustForRows, translateFormula } from './adjust.js'

const own = { sheet: 'Plan1', own: true }
const other = { sheet: 'Dados', own: false }

describe('copying the formula', () => {
  it('shifts relative references', () => {
    expect(translateFormula('=A1+B2', 1, 0)).toBe('=A2+B3')
  })

  it('leaves absolute ones alone, which is what $ is for', () => {
    expect(translateFormula('=$A$1+A1', 2, 0)).toBe('=$A$1+A3')
  })

  it('locks only the marked axis', () => {
    expect(translateFormula('=A$1+$A1', 1, 1)).toBe('=B$1+$A2')
  })

  it('leaving the sheet becomes #REF!', () => {
    // Copying =A1 to column A has nowhere to point.
    expect(translateFormula('=A1', 0, -1)).toBe('=#REF!')
  })

  it('moves the whole range along', () => {
    expect(translateFormula('=SOMA(A1:A3)', 1, 0)).toBe('=SOMA(A2:A4)')
  })

  it('preserves spaces, case and the rest of the text', () => {
    // The rewrite works on tokens, not on the tree: the formula comes back as the user typed it,
    // except for the references that had to change.
    expect(translateFormula('= soma( a1 ; "a1" )', 1, 0)).toBe('= soma( A2 ; "a1" )')
  })
})

describe('inserting a row', () => {
  it('pushes the reference that was below', () => {
    expect(adjustForRows('=A5', 2, 1, own)).toBe('=A6')
  })

  it('leaves what is above alone', () => {
    expect(adjustForRows('=A1', 2, 1, own)).toBe('=A1')
  })

  it('also moves absolute ones, because the cell really moved', () => {
    // Here $ does not protect: what was at $A$5 is now at $A$6.
    expect(adjustForRows('=$A$5', 2, 1, own)).toBe('=$A$6')
  })

  it('stretches the range when the row enters the middle of it', () => {
    expect(adjustForRows('=SOMA(A1:A5)', 2, 1, own)).toBe('=SOMA(A1:A6)')
  })

  it('pushes the whole range when the row enters before it', () => {
    expect(adjustForRows('=SOMA(A3:A5)', 0, 1, own)).toBe('=SOMA(A4:A6)')
  })
})

describe('deleting a row', () => {
  it('pulls up the reference below', () => {
    expect(adjustForRows('=A5', 0, -1, own)).toBe('=A4')
  })

  it('a reference to the deleted row becomes #REF!', () => {
    expect(adjustForRows('=A3', 2, -1, own)).toBe('=#REF!')
  })

  it('shrinks the range instead of destroying it', () => {
    // The easiest part to get wrong: treating the ends separately would turn the first into #REF!
    // and kill the formula over a deletion Excel absorbs without complaint.
    expect(adjustForRows('=SOMA(A1:A5)', 0, -3, own)).toBe('=SOMA(A1:A2)')
  })

  it('shrinks from the end too', () => {
    expect(adjustForRows('=SOMA(A1:A5)', 3, -2, own)).toBe('=SOMA(A1:A3)')
  })

  it('becomes #REF! only when the whole range disappears', () => {
    expect(adjustForRows('=SOMA(A1:A5)', 0, -5, own)).toBe('=SOMA(#REF!:#REF!)')
  })
})

describe('columns', () => {
  it('inserting pushes to the right', () => {
    expect(adjustForColumns('=C1', 1, 1, own)).toBe('=D1')
  })

  it('deleting pulls to the left', () => {
    expect(adjustForColumns('=C1', 0, -1, own)).toBe('=B1')
  })

  it('shrinks the range', () => {
    expect(adjustForColumns('=SOMA(A1:E1)', 0, -2, own)).toBe('=SOMA(A1:C1)')
  })
})

describe('across sheets', () => {
  it('a formula on the sheet itself adjusts the unqualified reference', () => {
    expect(adjustForRows('=A5', 0, 1, own)).toBe('=A6')
  })

  it('a formula on another sheet does not adjust the unqualified reference', () => {
    // =A1 on the "Resumo" sheet points to "Resumo", not to "Dados".
    expect(adjustForRows('=A5', 0, 1, other)).toBe('=A5')
  })

  it('but adjusts the one that names the changed sheet', () => {
    expect(adjustForRows('=Dados!A5', 0, 1, other)).toBe('=Dados!A6')
  })

  it('does not confuse it with any other sheet', () => {
    expect(adjustForRows('=Outra!A5', 0, 1, other)).toBe('=Outra!A5')
  })

  it('keeps the name only on the first end of the range', () => {
    expect(adjustForRows('=SOMA(Dados!A1:A5)', 0, 1, other)).toBe('=SOMA(Dados!A2:A6)')
  })

  it('ignores case in the sheet name', () => {
    expect(adjustForRows('=dados!A5', 0, 1, other)).toBe('=dados!A6')
  })
})

describe('what must not be touched', () => {
  it('text that looks like a reference', () => {
    expect(adjustForRows('=SE(A1="B2";"C3";A5)', 0, 1, own)).toBe('=SE(A2="B2";"C3";A6)')
  })

  it('function name with letters and digits', () => {
    expect(translateFormula('=HOJE()+A1', 1, 0)).toBe('=HOJE()+A2')
  })

  it('a formula that cannot even be read comes back intact', () => {
    expect(adjustForRows('=SOMA("', 0, 1, own)).toBe('=SOMA("')
  })

  it('a formula without the equals sign works too', () => {
    expect(adjustForRows('A5', 0, 1, own)).toBe('A6')
  })
})
