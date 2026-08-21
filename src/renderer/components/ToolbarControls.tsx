import { Icon, type IconName } from './Icon.js'

interface ToolbarButtonProps {
  readonly icon: IconName
  readonly label: string
  readonly onClick: () => void
  /**
   * Só para botões que ligam e desligam algo.
   *
   * Deixado de fora quando o botão apenas executa — aí não há estado a mostrar,
   * e anunciar `aria-pressed="false"` faria o leitor de tela chamar de
   * interruptor desligado o que na verdade é um comando.
   */
  readonly active?: boolean
  readonly disabled?: boolean
  readonly shortcut?: string
}

export function ToolbarButton({
  icon,
  label,
  onClick,
  active,
  disabled = false,
  shortcut,
}: ToolbarButtonProps): React.JSX.Element {
  return (
    <button
      type="button"
      className={active === true ? 'tbtn tbtn--active' : 'tbtn'}
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
  readonly icon: IconName
  readonly label: string
  readonly value: string
  readonly onChange: (value: string) => void
  readonly onClear: () => void
}

/**
 * Escolher cor: ícone do que recebe a cor, sobre uma barra com a cor atual.
 *
 * É o formato do Word, do Writer e do Google Docs — e resolve a pergunta que um
 * quadradinho colorido sozinho não responde: se aquela cor é do texto ou do
 * fundo. O seletor nativo continua ali, invisível por cima: ele é quem abre o
 * diálogo de cores do sistema, mas sua aparência muda de sistema para sistema e
 * não combina com o resto da barra.
 */
export function ColorControl({
  icon,
  label,
  value,
  onChange,
  onClear,
}: ColorControlProps): React.JSX.Element {
  const clearLabel = `Remover ${label.toLowerCase()}`

  return (
    <span className="tcolor">
      <span className="tcolor__pick" title={label}>
        <Icon name={icon} />
        <span className="tcolor__bar" style={{ background: value }} />
        <input
          type="color"
          className="tcolor__input"
          value={value}
          aria-label={label}
          onChange={(event) => onChange(event.target.value)}
        />
      </span>
      <button
        type="button"
        className="tcolor__clear"
        onClick={onClear}
        title={clearLabel}
        aria-label={clearLabel}
        onMouseDown={(event) => event.preventDefault()}
      >
        ✕
      </button>
    </span>
  )
}

/**
 * Grupo de botões afins.
 *
 * Além de nomear o conjunto para quem usa leitor de tela, é o que faz a barra
 * quebrar em linha inteira quando a janela aperta: o grupo não se parte, então
 * "alinhar à direita" nunca aparece sozinho no começo da segunda linha.
 */
export function ToolbarGroup({
  label,
  children,
}: {
  readonly label: string
  readonly children: React.ReactNode
}): React.JSX.Element {
  return (
    <div className="tgroup" role="group" aria-label={label}>
      {children}
    </div>
  )
}

export function ToolbarSeparator(): React.JSX.Element {
  return <span className="tsep" role="separator" />
}
