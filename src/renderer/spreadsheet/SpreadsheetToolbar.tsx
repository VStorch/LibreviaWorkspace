import {
  CellFormat,
  getCell,
  type BorderSide,
  type CellStyle,
  type Sheet,
} from '@services/spreadsheet/model.js'
import { applyBorders, applyStyle, toggleStyle, type Range } from '@services/spreadsheet/edit.js'
import type { MessageKey } from '@shared/i18n/index.js'
import {
  ColorControl,
  ToolbarButton,
  ToolbarGroup,
  ToolbarSelect,
  ToolbarSeparator,
} from '../components/ToolbarControls.js'
import { useT } from '../i18n.js'

function numberFormats(t: (key: MessageKey) => string) {
  return [
    { value: CellFormat.General, label: t('spreadsheet.format.general') },
    { value: CellFormat.Number, label: t('spreadsheet.format.number') },
    { value: CellFormat.Currency, label: t('spreadsheet.format.currency') },
    { value: CellFormat.Percent, label: t('spreadsheet.format.percent') },
    { value: CellFormat.Date, label: t('spreadsheet.format.date') },
    { value: CellFormat.Text, label: t('spreadsheet.format.text') },
  ]
}

const ALL_SIDES: readonly BorderSide[] = ['top', 'right', 'bottom', 'left']

/**
 * Barra de ferramentas da planilha.
 *
 * Age sobre a **seleção**, não sobre uma célula: formatar uma coluna inteira é
 * o uso normal, e obrigar clique a clique seria inútil na prática.
 *
 * Os botões mostram o estado da seleção — negrito fica marcado quando *toda* a
 * seleção está em negrito, que é a mesma regra que o clique aplica. Um botão
 * que acende com a seleção mista mentiria sobre o que o próximo clique faz.
 *
 * Usa os mesmos controles da barra do documento: numa suíte, negrito precisa
 * ser o mesmo botão nos dois editores.
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
  const t = useT()
  const style = styleOfSelection(sheet, range)

  const toggle = (key: 'bold' | 'italic' | 'underline') => () => onChange(toggleStyle(sheet, range, key))
  const set = (change: Partial<CellStyle>) => () => onChange(applyStyle(sheet, range, change))

  return (
    <div className="toolbar" role="toolbar" aria-label={t('spreadsheet.toolbar.label')}>
      {/* A referência da seleção mora na barra de fórmulas, logo abaixo, que é
          onde o Excel a põe — repeti-la aqui seria ruído. */}
      <ToolbarGroup label={t('spreadsheet.toolbar.textFormat')}>
        <ToolbarButton
          icon="bold"
          label={t('spreadsheet.toolbar.bold')}
          shortcut="Ctrl+B"
          active={style.bold === true}
          onClick={toggle('bold')}
        />
        <ToolbarButton
          icon="italic"
          label={t('spreadsheet.toolbar.italic')}
          shortcut="Ctrl+I"
          active={style.italic === true}
          onClick={toggle('italic')}
        />
        <ToolbarButton
          icon="underline"
          label={t('spreadsheet.toolbar.underline')}
          shortcut="Ctrl+U"
          active={style.underline === true}
          onClick={toggle('underline')}
        />
        <ColorControl
          icon="text-color"
          label={t('spreadsheet.toolbar.textColor')}
          value={style.color ?? '#000000'}
          onChange={(value) => onChange(applyStyle(sheet, range, { color: value }))}
          onClear={() => onChange(applyStyle(sheet, range, { color: undefined }))}
        />
        <ColorControl
          icon="fill-color"
          label={t('spreadsheet.toolbar.fillColor')}
          value={style.background ?? '#ffffff'}
          onChange={(value) => onChange(applyStyle(sheet, range, { background: value }))}
          onClear={() => onChange(applyStyle(sheet, range, { background: undefined }))}
        />
      </ToolbarGroup>

      <ToolbarSeparator />

      <ToolbarGroup label={t('spreadsheet.toolbar.alignment')}>
        <ToolbarButton
          icon="align-left"
          label={t('spreadsheet.toolbar.alignLeft')}
          active={style.align === 'left'}
          onClick={set({ align: 'left' })}
        />
        <ToolbarButton
          icon="align-center"
          label={t('spreadsheet.toolbar.alignCenter')}
          active={style.align === 'center'}
          onClick={set({ align: 'center' })}
        />
        <ToolbarButton
          icon="align-right"
          label={t('spreadsheet.toolbar.alignRight')}
          active={style.align === 'right'}
          onClick={set({ align: 'right' })}
        />
      </ToolbarGroup>

      <ToolbarSeparator />

      <ToolbarGroup label={t('spreadsheet.toolbar.numberGroup')}>
        <ToolbarSelect
          label={t('spreadsheet.toolbar.numberFormat')}
          value={style.format ?? CellFormat.General}
          options={numberFormats(t)}
          onChange={(value) => onChange(applyStyle(sheet, range, { format: value }))}
          width={124}
        />
        <ToolbarButton
          icon="decimal-less"
          label={t('spreadsheet.toolbar.decreaseDecimals')}
          onClick={set({ decimals: Math.max(0, (style.decimals ?? 2) - 1) })}
        />
        <ToolbarButton
          icon="decimal-more"
          label={t('spreadsheet.toolbar.increaseDecimals')}
          onClick={set({ decimals: Math.min(10, (style.decimals ?? 0) + 1) })}
        />
      </ToolbarGroup>

      <ToolbarSeparator />

      <ToolbarGroup label={t('spreadsheet.toolbar.borders')}>
        <ToolbarButton
          icon="borders-all"
          label={t('spreadsheet.toolbar.allBorders')}
          onClick={() => onChange(applyBorders(sheet, range, ALL_SIDES))}
        />
        <ToolbarButton
          icon="borders-none"
          label={t('spreadsheet.toolbar.noBorders')}
          onClick={() => onChange(applyBorders(sheet, range, []))}
        />
      </ToolbarGroup>

      <ToolbarSeparator />

      <ToolbarGroup label={t('spreadsheet.toolbar.panes')}>
        {/* Congelar usa a seleção como referência: tudo acima e à esquerda dela
            fica preso, que é como o Excel e o Google Sheets fazem. */}
        <ToolbarButton
          icon="freeze"
          label={t('spreadsheet.toolbar.freeze')}
          onClick={() => onChange({ ...sheet, frozenRows: range.fromRow, frozenColumns: range.fromColumn })}
        />
        <ToolbarButton
          icon="unfreeze"
          label={t('spreadsheet.toolbar.unfreeze')}
          disabled={sheet.frozenRows === 0 && sheet.frozenColumns === 0}
          onClick={() => onChange({ ...sheet, frozenRows: 0, frozenColumns: 0 })}
        />
      </ToolbarGroup>
    </div>
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
