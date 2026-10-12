import { useMemo } from 'react'
import type { WorkbookModel } from '@services/spreadsheet/model.js'
import { workbookStatistics } from '@services/spreadsheet/statistics.js'
import { FilePropertiesDialog, type StatisticRow } from '../components/FilePropertiesDialog.js'
import { useLanguage } from '../i18n.js'

/** The file's properties with the spreadsheet's counts. */
export function WorkbookPropertiesDialog({
  workbook,
  onClose,
}: {
  readonly workbook: WorkbookModel
  readonly onClose: () => void
}): React.JSX.Element {
  const language = useLanguage()
  const counts = useMemo(() => workbookStatistics(workbook), [workbook])
  const number = (value: number): string => value.toLocaleString(language)

  const statistics: readonly StatisticRow[] = [
    ['spreadsheet.properties.sheets', number(counts.sheets)],
    ['spreadsheet.properties.filledCells', number(counts.filledCells)],
    ['spreadsheet.properties.formulas', number(counts.formulas)],
  ]

  return (
    <FilePropertiesDialog title="spreadsheet.properties.dialog" statistics={statistics} onClose={onClose} />
  )
}
