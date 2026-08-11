import { Icon, type IconName } from './Icon.js'

interface ToolbarButtonProps {
  readonly icon: IconName
  readonly label: string
  readonly onClick: () => void
  readonly active?: boolean
  readonly disabled?: boolean
  readonly shortcut?: string
}

export function ToolbarButton({
  icon,
  label,
  onClick,
  active = false,
  disabled = false,
  shortcut,
}: ToolbarButtonProps): React.JSX.Element {
  return (
    <button
      type="button"
      className={active ? 'tbtn tbtn--active' : 'tbtn'}
      onClick={onClick}
      disabled={disabled}
      title={shortcut === undefined ? label : `${label} (${shortcut})`}
      aria-label={label}
      aria-pressed={active}
      // Sem isto, clicar num botão tira o foco do editor e a seleção some
      // antes de o comando rodar.
      onMouseDown={(event) => event.preventDefault()}
    >
      <Icon name={icon} />
    </button>
  )
}

interface ToolbarSelectProps<T extends string> {
  readonly label: string
  readonly value: T
  readonly options: readonly { readonly value: T; readonly label: string }[]
  readonly onChange: (value: T) => void
  readonly width?: number
}

export function ToolbarSelect<T extends string>({
  label,
  value,
  options,
  onChange,
  width,
}: ToolbarSelectProps<T>): React.JSX.Element {
  return (
    <select
      className="tselect"
      value={value}
      aria-label={label}
      title={label}
      style={width === undefined ? undefined : { width: `${width}px` }}
      onChange={(event) => onChange(event.target.value as T)}
      onMouseDown={(event) => event.stopPropagation()}
    >
      {options.map((option) => (
        <option key={option.value} value={option.value}>
          {option.label}
        </option>
      ))}
    </select>
  )
}

interface ColorControlProps {
  readonly label: string
  readonly value: string
  readonly onChange: (value: string) => void
  readonly onClear: () => void
}

export function ColorControl({ label, value, onChange, onClear }: ColorControlProps): React.JSX.Element {
  return (
    <span className="tcolor" title={label}>
      <input
        type="color"
        className="tcolor__input"
        value={value}
        aria-label={label}
        onChange={(event) => onChange(event.target.value)}
      />
      <button
        type="button"
        className="tcolor__clear"
        onClick={onClear}
        title={`Remover ${label.toLowerCase()}`}
      >
        ✕
      </button>
    </span>
  )
}

export function ToolbarSeparator(): React.JSX.Element {
  return <span className="tsep" role="separator" />
}
