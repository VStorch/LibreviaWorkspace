import {
  CellFormat,
  getCell,
  type BorderSide,
  type CellStyle,
  type Sheet,
} from '@services/spreadsheet/model.js'
import {
  applyBorders,
  applyStyle,
  describeRange,
  toggleStyle,
  type Range,
} from '@services/spreadsheet/edit.js'

/**
 * Barra de ferramentas da planilha.
 *
 * Age sobre a **seleção**, não sobre uma célula: formatar uma coluna inteira é
 * o uso normal, e obrigar clique a clique seria inútil na prática.
 *
 * Os botões mostram o estado da seleção — negrito fica marcado quando *toda* a
 * seleção está em negrito, que é a mesma regra que o clique aplica. Um botão
 * que acende com a seleção mista mentiria sobre o que o próximo clique faz.
 */
export function SpreadsheetToolbar({
  sheet,
  range,
  onChange,
}: {
  sheet: Sheet
  range: Range
  onChange: (sheet: Sheet) => void
}): React.JSX.Element {
  const style = styleOfSelection(sheet, range)

  const toggle = (key: 'bold' | 'italic' | 'underline') => () => onChange(toggleStyle(sheet, range, key))
  const set = (change: Partial<CellStyle>) => () => onChange(applyStyle(sheet, range, change))

  return (
    <div className="sheet-toolbar" role="toolbar" aria-label="Formatação da planilha">
      <span className="sheet-toolbar__ref" title="Seleção atual">
        {describeRange(range)}
      </span>

      <span className="sheet-toolbar__sep" />

      <Toggle label="N" title="Negrito (Ctrl+B)" active={style.bold === true} onClick={toggle('bold')} bold />
      <Toggle
        label="I"
        title="Itálico (Ctrl+I)"
        active={style.italic === true}
        onClick={toggle('italic')}
        italic
      />
      <Toggle
        label="S"
        title="Sublinhado (Ctrl+U)"
        active={style.underline === true}
        onClick={toggle('underline')}
        underline
      />

      <span className="sheet-toolbar__sep" />

      <label className="sheet-toolbar__color" title="Cor do texto">
        <span aria-hidden="true">A</span>
        <input
          type="color"
          value={style.color ?? '#000000'}
          onChange={(event) => onChange(applyStyle(sheet, range, { color: event.target.value }))}
        />
      </label>
      <label className="sheet-toolbar__color" title="Cor de fundo">
        <span aria-hidden="true">▧</span>
        <input
          type="color"
          value={style.background ?? '#ffffff'}
          onChange={(event) => onChange(applyStyle(sheet, range, { background: event.target.value }))}
        />
      </label>
      <button
        type="button"
        className="sheet-toolbar__button"
        title="Remover cores"
        onClick={set({ color: undefined, background: undefined })}
      >
        ⌫
      </button>

      <span className="sheet-toolbar__sep" />

      <Toggle
        label="⇤"
        title="Alinhar à esquerda"
        active={style.align === 'left'}
        onClick={set({ align: 'left' })}
      />
      <Toggle
        label="↔"
        title="Centralizar"
        active={style.align === 'center'}
        onClick={set({ align: 'center' })}
      />
      <Toggle
        label="⇥"
        title="Alinhar à direita"
        active={style.align === 'right'}
        onClick={set({ align: 'right' })}
      />

      <span className="sheet-toolbar__sep" />

      <select
        className="sheet-toolbar__select"
        title="Formato do número"
        value={style.format ?? CellFormat.General}
        onChange={(event) => onChange(applyStyle(sheet, range, { format: event.target.value as CellFormat }))}
      >
        <option value={CellFormat.General}>Geral</option>
        <option value={CellFormat.Number}>Número</option>
        <option value={CellFormat.Currency}>Moeda</option>
        <option value={CellFormat.Percent}>Percentual</option>
        <option value={CellFormat.Date}>Data</option>
        <option value={CellFormat.Text}>Texto</option>
      </select>

      <button
        type="button"
        className="sheet-toolbar__button"
        title="Menos casas decimais"
        onClick={set({ decimals: Math.max(0, (style.decimals ?? 2) - 1) })}
      >
        ,0←
      </button>
      <button
        type="button"
        className="sheet-toolbar__button"
        title="Mais casas decimais"
        onClick={set({ decimals: Math.min(10, (style.decimals ?? 0) + 1) })}
      >
        ,00→
      </button>

      <span className="sheet-toolbar__sep" />

      <button
        type="button"
        className="sheet-toolbar__button"
        title="Bordas em volta"
        onClick={() => onChange(applyBorders(sheet, range, ALL_SIDES))}
      >
        ▣
      </button>
      <button
        type="button"
        className="sheet-toolbar__button"
        title="Sem bordas"
        onClick={() => onChange(applyBorders(sheet, range, []))}
      >
        ▢
      </button>

      <span className="sheet-toolbar__sep" />

      {/* Congelar usa a seleção como referência: tudo acima e à esquerda dela
          fica preso, que é como o Excel e o Google Sheets fazem. */}
      <button
        type="button"
        className="sheet-toolbar__button"
        title="Congelar até a seleção"
        onClick={() => onChange({ ...sheet, frozenRows: range.fromRow, frozenColumns: range.fromColumn })}
      >
        ❄
      </button>
      <button
        type="button"
        className="sheet-toolbar__button"
        title="Descongelar"
        onClick={() => onChange({ ...sheet, frozenRows: 0, frozenColumns: 0 })}
        disabled={sheet.frozenRows === 0 && sheet.frozenColumns === 0}
      >
        ☀
      </button>
    </div>
  )
}

const ALL_SIDES: readonly BorderSide[] = ['top', 'right', 'bottom', 'left']

function Toggle({
  label,
  title,
  active,
  onClick,
  bold,
  italic,
  underline,
}: {
  label: string
  title: string
  active: boolean
  onClick: () => void
  bold?: boolean
  italic?: boolean
  underline?: boolean
}): React.JSX.Element {
  return (
    <button
      type="button"
      className={active ? 'sheet-toolbar__button sheet-toolbar__button--active' : 'sheet-toolbar__button'}
      title={title}
      aria-pressed={active}
      onClick={onClick}
      style={{
        fontWeight: bold === true ? 700 : undefined,
        fontStyle: italic === true ? 'italic' : undefined,
        textDecoration: underline === true ? 'underline' : undefined,
      }}
    >
      {label}
    </button>
  )
}

/**
 * O estilo comum a toda a seleção.
 *
 * Só devolve um atributo quando **todas** as células concordam. É o que faz o
 * botão de negrito acender apenas quando a seleção inteira está em negrito — o
 * mesmo critério que o clique usa para decidir entre ligar e desligar.
 */
function styleOfSelection(sheet: Sheet, range: Range): CellStyle {
  const first = getCell(sheet, range.fromRow, range.fromColumn)?.style ?? {}
  const common: Record<string, unknown> = { ...first }

  for (let row = range.fromRow; row <= range.toRow; row++) {
    for (let column = range.fromColumn; column <= range.toColumn; column++) {
      const style = (getCell(sheet, row, column)?.style ?? {}) as Record<string, unknown>
      for (const key of Object.keys(common)) {
        if (style[key] !== common[key]) delete common[key]
      }
    }
  }

  return common as CellStyle
}
