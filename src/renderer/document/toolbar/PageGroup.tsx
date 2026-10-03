import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { setPreference, usePreferences } from '../../state/preferences.js'
import { useWorkspace } from '../../state/workspace.js'

interface PageGroupProps {
  readonly onOpenFind: () => void
  readonly onOpenPageSetup: () => void
}

export function PageGroup({ onOpenFind, onOpenPageSetup }: PageGroupProps): React.JSX.Element {
  const t = useT()
  const printPreview = useWorkspace((state) => state.printPreview)
  // Do main: a barra e o menu "Exibir" mudam a mesma chave.
  const invisibleCharacters = usePreferences((state) => state.preferences.invisibleCharacters)

  return (
    <ToolbarGroup label={t('document.pageGroup.group')}>
      <ToolbarButton
        icon="formatting-marks"
        label={t('menu.view.formattingMarks')}
        shortcut={shortcutHintOf(SHORTCUTS.formattingMarks)}
        active={invisibleCharacters}
        onClick={() => void setPreference({ invisibleCharacters: !invisibleCharacters })}
      />
      <ToolbarButton
        icon="search"
        label={t('document.pageGroup.findReplace')}
        shortcut={shortcutHintOf(SHORTCUTS.findReplace)}
        onClick={onOpenFind}
      />
      <ToolbarButton icon="page-setup" label={t('document.pageSetup.title')} onClick={onOpenPageSetup} />
      <ToolbarButton
        icon="print-preview"
        label={t('menu.file.printPreview')}
        onClick={() => void printPreview()}
      />
    </ToolbarGroup>
  )
}
