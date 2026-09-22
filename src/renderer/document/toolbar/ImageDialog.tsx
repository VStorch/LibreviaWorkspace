import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useT } from '../../i18n.js'
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
  const t = useT()
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
      aria-label={t('document.imageDialog.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        if (event.key === 'Enter') apply()
      }}
    >
      <label className="popover__field">
        <span>{t('document.imageDialog.altText')}</span>
        <input
          type="text"
          aria-label={t('document.imageDialog.altText')}
          maxLength={MAX_ALT_LENGTH}
          value={alt}
          autoFocus
          onChange={(event) => setAlt(event.target.value)}
        />
      </label>

      <label className="popover__field">
        <span>{t('document.imageDialog.alignment')}</span>
        <select
          aria-label={t('document.imageDialog.alignmentLabel')}
          value={align}
          onChange={(event) => setAlign(event.target.value)}
        >
          <option value="">{t('document.imageDialog.alignSameAsParagraph')}</option>
          <option value="left">{t('document.imageDialog.alignLeft')}</option>
          <option value="center">{t('document.imageDialog.alignCenter')}</option>
          <option value="right">{t('document.imageDialog.alignRight')}</option>
        </select>
      </label>

      <p className="popover__hint">{t('document.imageDialog.hint')}</p>

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onMouseDown={keepFocus} onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onMouseDown={keepFocus} onClick={apply}>
          {t('document.common.apply')}
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
