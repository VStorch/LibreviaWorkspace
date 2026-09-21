import { SHORTCUTS, shortcutHintOf } from '@shared/shortcuts.js'
import { ToolbarButton, ToolbarGroup } from '../../components/ToolbarControls.js'
import { setPreference, usePreferences } from '../../state/preferences.js'
import { useWorkspace } from '../../state/workspace.js'

interface PageGroupProps {
  readonly onOpenFind: () => void
  readonly onOpenPageSetup: () => void
}

/** A página e o documento inteiro: nada aqui depende da seleção. */
export function PageGroup({ onOpenFind, onOpenPageSetup }: PageGroupProps): React.JSX.Element {
  const printPreview = useWorkspace((state) => state.printPreview)
  // Do main, que é o dono da preferência: clicar aqui e clicar no item do menu
  // "Exibir" mudam a mesma chave, e os dois ficam marcados juntos.
  const invisibleCharacters = usePreferences((state) => state.preferences.invisibleCharacters)

  return (
    <ToolbarGroup label="Página">
      <ToolbarButton
        icon="formatting-marks"
        label="Marcas de formatação"
        shortcut={shortcutHintOf(SHORTCUTS.formattingMarks)}
        active={invisibleCharacters}
        onClick={() => void setPreference({ invisibleCharacters: !invisibleCharacters })}
      />
      <ToolbarButton
        icon="search"
        label="Localizar e substituir"
        shortcut={shortcutHintOf(SHORTCUTS.findReplace)}
        onClick={onOpenFind}
      />
      <ToolbarButton icon="page-setup" label="Configuração de página" onClick={onOpenPageSetup} />
      {/* Como o editor não pagina ao vivo (§6.3 do plano), a prévia é o que
        responde "onde as páginas quebram" — e por isso fica à mão. */}
      <ToolbarButton icon="print-preview" label="Visualizar impressão" onClick={() => void printPreview()} />
    </ToolbarGroup>
  )
}
