import { useState } from 'react'
import type { WorkbookModel } from '@services/spreadsheet/model.js'
import { useT } from '../i18n.js'

/**
 * Abas da pasta de trabalho.
 *
 * Renomear é por duplo clique, como em toda planilha — e o nome é validado:
 * duas abas com o mesmo nome quebrariam a referência entre abas que o motor de
 * fórmulas da Fase 6 vai precisar.
 */
export function SheetTabs({
  workbook,
  onSelect,
  onAdd,
  onRename,
  onRemove,
}: {
  workbook: WorkbookModel
  onSelect: (index: number) => void
  onAdd: () => void
  onRename: (index: number, name: string) => void
  onRemove: (index: number) => void
}): React.JSX.Element {
  const t = useT()
  const [editing, setEditing] = useState<number | null>(null)

  return (
    <div className="tabs" role="tablist">
      {workbook.sheets.map((sheet, index) => (
        <div
          key={index}
          className={index === workbook.activeSheet ? 'tab tab--active' : 'tab'}
          role="tab"
          aria-selected={index === workbook.activeSheet}
        >
          {editing === index ? (
            <input
              className="tab__input"
              defaultValue={sheet.name}
              autoFocus
              onBlur={(event) => {
                onRename(index, event.target.value)
                setEditing(null)
              }}
              onKeyDown={(event) => {
                if (event.key === 'Enter') event.currentTarget.blur()
                // Esc desiste sem gravar: renomear por engano é fácil de fazer.
                if (event.key === 'Escape') setEditing(null)
              }}
            />
          ) : (
            <button
              type="button"
              className="tab__name"
              onClick={() => onSelect(index)}
              onDoubleClick={() => setEditing(index)}
              title={t('spreadsheet.sheetTabs.doubleClickToRename')}
            >
              {sheet.name}
            </button>
          )}

          {/* A última aba não pode ser removida: uma pasta sem planilha nenhuma
              não é um estado que o modelo aceite. */}
          {workbook.sheets.length > 1 && editing !== index && (
            <button
              type="button"
              className="tab__close"
              onClick={() => onRemove(index)}
              aria-label={t('spreadsheet.sheetTabs.deleteSheet', { name: sheet.name })}
              title={t('spreadsheet.sheetTabs.deleteSheet', { name: sheet.name })}
            >
              ✕
            </button>
          )}
        </div>
      ))}

      <button type="button" className="tabs__add" onClick={onAdd} aria-label={t('spreadsheet.sheetTabs.addSheet')}>
        +
      </button>
    </div>
  )
}
