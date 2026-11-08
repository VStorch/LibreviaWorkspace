import type { CommandId } from '@shared/commands/registry.js'
import { ContextMenuItem } from '../components/ContextMenu.js'
import type { IconName } from '../components/Icon.js'
import { ToolbarButton } from '../components/ToolbarControls.js'
import { useCommand } from './useCommand.js'

/** The label stays the toolbar's: the menu's says less, out of its menu. */
export function CommandButton({
  command,
  icon,
  label,
}: {
  command: CommandId
  icon: IconName
  label: string
}): React.JSX.Element {
  const { enabled, shortcut, run } = useCommand(command)
  return (
    <ToolbarButton
      icon={icon}
      label={label}
      {...(shortcut === undefined ? {} : { shortcut })}
      disabled={!enabled}
      onClick={run}
    />
  )
}

export function CommandMenuItem({
  command,
  disabled = false,
  onClose,
  children,
}: {
  command: CommandId
  disabled?: boolean
  onClose: () => void
  children: React.ReactNode
}): React.JSX.Element {
  const { enabled, run } = useCommand(command)
  return (
    <ContextMenuItem
      disabled={disabled || !enabled}
      onClick={() => {
        onClose()
        run()
      }}
    >
      {children}
    </ContextMenuItem>
  )
}
