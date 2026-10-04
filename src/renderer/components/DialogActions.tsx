import type { ReactNode } from 'react'
import { useT } from '../i18n.js'

/** O clique não tira o foco do editor: o comando age sobre a seleção que havia. */
const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

/** Cancelar e confirmar, à direita; `children` vai à esquerda. */
export function DialogActions({
  confirmLabel,
  onConfirm,
  onCancel,
  disabled = false,
  keepEditorFocus = false,
  children,
}: {
  readonly confirmLabel: string
  readonly onConfirm: () => void
  readonly onCancel: () => void
  readonly disabled?: boolean
  readonly keepEditorFocus?: boolean
  readonly children?: ReactNode
}): React.JSX.Element {
  const t = useT()
  const onMouseDown = keepEditorFocus ? keepFocus : undefined
  return (
    <div className="popover__actions">
      {children}
      <span className="popover__spacer" />
      <button type="button" className="btn" onMouseDown={onMouseDown} onClick={onCancel}>
        {t('document.common.cancel')}
      </button>
      <button
        type="button"
        className="btn btn--primary"
        onMouseDown={onMouseDown}
        onClick={onConfirm}
        disabled={disabled}
      >
        {confirmLabel}
      </button>
    </div>
  )
}
