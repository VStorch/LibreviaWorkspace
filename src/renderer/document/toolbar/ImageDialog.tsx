import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { applyImageProperties, imageAt } from '../extensions/document-image.js'

/** O tamanho do campo de texto alternativo que o `w:docPr/@descr` aceita sem exagero. */
const MAX_ALT_LENGTH = 300

/**
 * Propriedades da imagem: texto alternativo e alinhamento.
 *
 * O texto alternativo é acessibilidade de verdade — vai para `wp:docPr/@descr`, é
 * o que um leitor de tela anuncia no lugar da imagem, e é o campo que o Word
 * chama de "Texto Alt". Até aqui o editor escrevia nele o **nome do arquivo**, que
 * é melhor que nada e não é uma descrição.
 *
 * O tamanho não está aqui de propósito: ele se resolve arrastando as alças da
 * própria imagem, que é onde se espera mexer nele.
 */
export function ImageDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const placed = imageAt(editor)
  const [alt, setAlt] = useState(() => {
    const value = placed?.node.attrs['alt']
    return typeof value === 'string' ? value : ''
  })
  const [align, setAlign] = useState(() => currentAlign(editor))

  const keepFocus = (event: React.MouseEvent): void => event.preventDefault()

  function apply(): void {
    const target = imageAt(editor)
    if (target !== null) {
      applyImageProperties(editor, target, { alt, align: align === '' ? null : align })
    }
    onClose()
    requestAnimationFrame(() => editor.commands.focus())
  }

  return (
    <div
      className="popover"
      role="dialog"
      aria-label="Propriedades da imagem"
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <label className="popover__field">
        <span>Texto alternativo</span>
        <input
          type="text"
          aria-label="Texto alternativo"
          maxLength={MAX_ALT_LENGTH}
          value={alt}
          autoFocus
          onChange={(event) => setAlt(event.target.value)}
        />
      </label>

      <label className="popover__field">
        <span>Alinhamento</span>
        <select
          aria-label="Alinhamento da imagem"
          value={align}
          onChange={(event) => setAlign(event.target.value)}
        >
          <option value="">Como o parágrafo</option>
          <option value="left">À esquerda</option>
          <option value="center">Centralizada</option>
          <option value="right">À direita</option>
        </select>
      </label>

      <p className="popover__hint">
        O texto alternativo é lido em voz alta no lugar da imagem, e vai no arquivo.
      </p>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          Cancelar
        </button>
        <button type="button" className="btn btn--primary" onMouseDown={keepFocus} onClick={apply}>
          Aplicar
        </button>
      </div>
    </div>
  )
}

/**
 * O alinhamento que a imagem já tem — do parágrafo, quando ela mora num, ou do
 * atributo dela, quando é um bloco solto.
 */
function currentAlign(editor: Editor): string {
  const placed = imageAt(editor)
  if (placed === null) return ''

  if (placed.paragraphPos !== null) {
    const paragraph = editor.state.doc.nodeAt(placed.paragraphPos)
    const align = paragraph?.attrs['textAlign']
    return typeof align === 'string' ? align : ''
  }

  const align = placed.node.attrs['align']
  return typeof align === 'string' ? align : ''
}
