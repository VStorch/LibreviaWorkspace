import { shortcutHintOf } from '@shared/shortcuts.js'
import { MenuCommand } from '@shared/types.js'
import { CommandButton } from '../../commands/CommandControls.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { useT } from '../../i18n.js'
import { setPreference, usePreferences, useShortcutTable } from '../../state/preferences.js'

export function PageGroup(): React.JSX.Element {
  const t = useT()
  // From main: the toolbar and the "View" menu change the same key.
  const invisibleCharacters = usePreferences((state) => state.preferences.invisibleCharacters)
  const shortcuts = useShortcutTable()

  return (
    <ToolbarGroup label={t('document.pageGroup.group')}>
      <ToolbarButton
        icon="formatting-marks"
        label={t('menu.view.formattingMarks')}
        shortcut={shortcutHintOf(shortcuts.formattingMarks)}
        active={invisibleCharacters}
        onClick={() => void setPreference({ invisibleCharacters: !invisibleCharacters })}
      />
      <CommandButton
        command={MenuCommand.FindReplace}
        icon="search"
        label={t('document.pageGroup.findReplace')}
      />
      <CommandButton
        command={MenuCommand.PageSetup}
        icon="page-setup"
        label={t('document.pageSetup.title')}
      />
      <CommandButton
        command={MenuCommand.PrintPreview}
        icon="print-preview"
        label={t('menu.file.printPreview')}
      />
    </ToolbarGroup>
  )
}
