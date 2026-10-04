import { useRef, useState } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import {
  StyleType,
  blockStyleOf,
  listedStyles,
  styleLabelOf,
  type StyleDefinition,
  type StyleSheet,
} from '@services/document/styles.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { StyleDialog, type StyleDialogMode } from './StyleDialog.js'
import { cycleFocus } from '../components/focus-trap.js'

const STYLE_PANEL_STOPS = 'select, ul[tabindex], input, button:not(:disabled), li[tabindex="0"]'

/**
 * Aplicar e limpar são transações do editor (`style-commands.ts`); modificar e
 * criar trocam a folha do store, que vai a `word/styles.xml` (`StyleWriter.cs`).
 * Nenhum estilo é excluído. No somente leitura, só consulta.
 */
export function StylesPanel({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const sheet = useWorkspace((state) => state.styles)
  const setStyles = useWorkspace((state) => state.setStyles)
  const readOnly = useWorkspace((state) => state.readOnly)
  const [selected, setSelected] = useState<string | null>(null)
  const [editing, setEditing] = useState<StyleDialogMode | null>(null)
  const language = useLanguage()
  const t = useT()
  const [filter, setFilter] = useState<'all' | StyleType>('all')
  const panel = useRef<HTMLDivElement>(null)

  const current = blockStyleOf(sheet, useCursorBlock(editor))
  const styles = listedStyles(sheet, language).filter((style) => filter === 'all' || style.type === filter)
  const chosen = selected === null ? undefined : sheet.styles[selected]

  if (editing !== null) {
    return (
      <StyleEditing
        editor={editor}
        sheet={sheet}
        mode={editing}
        onCancel={() => setEditing(null)}
        onDone={(next, id) => {
          setStyles(next)
          setSelected(id)
          setEditing(null)
        }}
      />
    )
  }

  return (
    <div
      ref={panel}
      className="popover"
      role="dialog"
      aria-label={t('document.styles.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
        else if (event.key === 'Tab') cycleFocus(panel.current, event, STYLE_PANEL_STOPS)
      }}
    >
      <p className="popover__hint">
        {current === null
          ? t('document.styles.noCurrent')
          : t('document.styles.current', { style: styleLabelOf(current, language) })}
      </p>

      <StyleFilter filter={filter} onFilter={setFilter} />

      <StyleList
        styles={styles}
        sheet={sheet}
        currentId={current?.id ?? null}
        selected={selected}
        onSelect={setSelected}
        onApply={(style) => {
          if (!readOnly) applyStyle(editor, style)
        }}
      />

      {readOnly ? (
        <p className="popover__hint">{t('document.styles.readOnly')}</p>
      ) : (
        chosen === undefined && <p className="popover__hint">{t('document.styles.select')}</p>
      )}

      <StyleActions
        editor={editor}
        readOnly={readOnly}
        chosen={chosen}
        currentId={current?.id ?? null}
        onEdit={setEditing}
        onClose={onClose}
      />
    </div>
  )
}

function StyleEditing(props: React.ComponentProps<typeof StyleDialog>): React.JSX.Element {
  const t = useT()
  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label={t('document.styles.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') props.onCancel()
      }}
    >
      <StyleDialog {...props} />
    </div>
  )
}

function applyStyle(editor: Editor, style: StyleDefinition): void {
  const chain = editor.chain().focus()
  if (style.type === StyleType.Character) chain.applyCharacterStyle(style.id).run()
  else chain.applyParagraphStyle(style.id).run()
}

/** Ao vivo: o cursor anda com o painel aberto. */
function useCursorBlock(editor: Editor): { type: string; styleId: string | null; level: number | null } {
  return useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const parent = current.state.selection.$from.parent
      const styleId = parent.attrs['styleId']
      const level = parent.attrs['level']
      return {
        type: parent.type.name,
        styleId: typeof styleId === 'string' ? styleId : null,
        level: typeof level === 'number' ? level : null,
      }
    },
  })
}

function StyleFilter({
  filter,
  onFilter,
}: {
  filter: 'all' | StyleType
  onFilter: (filter: 'all' | StyleType) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__row">
      <label className="popover__field">
        <span>{t('document.styles.show')}</span>
        <select
          aria-label={t('document.styles.show')}
          value={filter}
          onChange={(event) => onFilter(event.target.value as 'all' | StyleType)}
        >
          <option value="all">{t('document.styles.all')}</option>
          <option value={StyleType.Paragraph}>{t('document.styles.paragraphFilter')}</option>
          <option value={StyleType.Character}>{t('document.styles.characterFilter')}</option>
        </select>
      </label>
    </div>
  )
}

function StyleList({
  styles,
  sheet,
  currentId,
  selected,
  onSelect,
  onApply,
}: {
  styles: readonly StyleDefinition[]
  sheet: StyleSheet
  currentId: string | null
  selected: string | null
  onSelect: (id: string) => void
  onApply: (style: StyleDefinition) => void
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  return (
    // Focalizável, para a lista rolar pelo teclado; clique duplo já aplica.
    <ul className="styles-list" tabIndex={0} aria-label={t('document.styleAndFont.documentStyles')}>
      {styles.length === 0 && <li className="styles-list__empty">{t('document.styles.empty')}</li>}
      {styles.map((style) => (
        <li
          key={style.id}
          className={[
            'styles-list__item',
            style.id === currentId ? 'styles-list__item--current' : '',
            style.id === selected ? 'styles-list__item--selected' : '',
          ]
            .filter((name) => name !== '')
            .join(' ')}
          aria-selected={style.id === selected}
          onClick={() => onSelect(style.id)}
          onDoubleClick={() => {
            onSelect(style.id)
            onApply(style)
          }}
          // A marca visual sozinha não diz nada a quem não vê a tela.
          aria-current={style.id === currentId ? 'true' : undefined}
        >
          <span className="styles-list__name">{styleLabelOf(style, language)}</span>
          <span className="styles-list__meta">{describe(style, sheet, language, t)}</span>
        </li>
      ))}
    </ul>
  )
}

function StyleActions({
  editor,
  readOnly,
  chosen,
  currentId,
  onEdit,
  onClose,
}: {
  editor: Editor
  readOnly: boolean
  chosen: StyleDefinition | undefined
  currentId: string | null
  onEdit: (mode: StyleDialogMode) => void
  onClose: () => void
}): React.JSX.Element {
  const t = useT()
  return (
    <div className="popover__actions">
      <button
        type="button"
        className="btn"
        disabled={readOnly || chosen === undefined}
        onClick={() => {
          if (chosen !== undefined && !readOnly) applyStyle(editor, chosen)
        }}
      >
        {t('document.styles.apply')}
      </button>
      <button
        type="button"
        className="btn"
        disabled={readOnly || chosen === undefined}
        onClick={() => chosen !== undefined && onEdit({ kind: 'modify', id: chosen.id })}
      >
        {t('document.styles.modify')}
      </button>
      <button
        type="button"
        className="btn"
        disabled={readOnly}
        onClick={() =>
          onEdit({ kind: 'create', basedOn: chosen?.type === StyleType.Paragraph ? chosen.id : currentId })
        }
      >
        {t('document.styles.create')}
      </button>
      <button
        type="button"
        className="btn"
        disabled={readOnly}
        onClick={() => editor.chain().focus().clearDirectFormatting().run()}
      >
        {t('document.styles.clearFormatting')}
      </button>
      <span className="popover__spacer" />
      <button type="button" className="btn btn--primary" autoFocus onClick={onClose}>
        {t('document.common.close')}
      </button>
    </div>
  )
}

/** Só o que o estilo declara, e de quem herda: "12 pt" num estilo que herda o tamanho afirmaria o que o documento não diz. */
function describe(
  style: StyleDefinition,
  sheet: StyleSheet,
  language: ReturnType<typeof useLanguage>,
  t: ReturnType<typeof useT>,
): string {
  const parts: string[] = [
    style.type === StyleType.Character ? t('document.styles.character') : t('document.styles.paragraph'),
  ]

  const font = style.character?.fontFamily
  if (font !== undefined) parts.push(font.split(',')[0]!.trim())
  if (style.character?.fontSize !== undefined) parts.push(style.character.fontSize)
  if (style.character?.bold === true) parts.push(t('document.styles.bold'))
  if (style.basedOn !== undefined) {
    // Pelo nome: num documento em português o id do pai é `Ttulo1`.
    const parent = sheet.styles[style.basedOn]
    parts.push(
      t('document.styles.basedOn', {
        style: parent === undefined ? style.basedOn : styleLabelOf(parent, language),
      }),
    )
  }

  return parts.join(' · ')
}
