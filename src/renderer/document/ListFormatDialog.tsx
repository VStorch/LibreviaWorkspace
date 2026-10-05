import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/react'
import type { MessageKey } from '@shared/i18n/index.js'
import type { DocumentNode } from '@services/document/model.js'
import {
  LIST_LEVELS,
  defaultLevels,
  numberLists,
  type LevelDef,
  type ListTreeReader,
} from '@services/document/list-numbering.js'
import { BULLET_MARKS, LIST_PRESETS, NUMBER_FORMATS, kindOfLevels } from '@services/document/list-presets.js'
import { useT } from '../i18n.js'
import { listEntries } from './extensions/list-numbering.js'
import { MAX_START_NUMBER } from '@shared/limits.js'

const FORMAT_LABELS: Record<(typeof NUMBER_FORMATS)[number], MessageKey> = {
  decimal: 'document.lists.fmtDecimal',
  decimalZero: 'document.lists.fmtDecimalZero',
  lowerLetter: 'document.lists.fmtLowerLetter',
  upperLetter: 'document.lists.fmtUpperLetter',
  lowerRoman: 'document.lists.fmtLowerRoman',
  upperRoman: 'document.lists.fmtUpperRoman',
  bullet: 'document.lists.fmtBullet',
  none: 'document.lists.fmtNone',
}

const JSON_READER: ListTreeReader<DocumentNode> = {
  typeOf: (node) => node.type,
  attrsOf: (node) => node.attrs ?? {},
  childrenOf: (node) => node.content ?? [],
}

/** Through the document's own numbering: a second `%1.%2.` grammar would drift from the paper. */
export function previewOf(levels: readonly LevelDef[]): string[] {
  const kind = kindOfLevels(levels)
  const numbering = { key: 'previa', levels }
  let nested: DocumentNode | null = null
  for (let level = 2; level >= 0; level--) {
    const item: DocumentNode = {
      type: 'listItem',
      content: [{ type: 'paragraph' }, ...(nested === null ? [] : [nested])],
    }
    nested = {
      type: kind,
      attrs: level === 0 ? { numbering } : { level },
      content: [item],
    }
  }
  return numberLists({ type: 'doc', content: [nested!] }, JSON_READER).labels
}

/** The levels start from the cursor's list, as in Word. */
export function ListFormatDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()

  const current = useMemo(() => {
    const { $from } = editor.state.selection
    const entries = listEntries(editor.state.doc)
    return entries
      .filter((entry) => entry.pos < $from.pos && entry.pos + entry.node.nodeSize > $from.pos)
      .at(-1)
  }, [editor])

  const [levels, setLevels] = useState<LevelDef[]>(() =>
    current === undefined
      ? defaultLevels('orderedList')
      : Array.from({ length: LIST_LEVELS }, (_, index) => ({
          ...(current.info.def.levels[index] ?? defaultLevels(current.info.kind)[index]!),
        })),
  )
  const [level, setLevel] = useState(current?.info.level ?? 0)

  const own = levels[level]!
  const preview = previewOf(levels)
  const bullet = own.fmt === 'bullet'
  const ordered = current !== undefined && current.info.kind === 'orderedList'

  const update = (change: Partial<LevelDef>): void =>
    setLevels((all) => all.map((entry, index) => (index === level ? { ...entry, ...change } : entry)))

  function changeFormat(fmt: string): void {
    // Switching from number to bullet switches the text: `%2.` on a bullet would draw "%2.".
    if (fmt === 'bullet') update({ fmt, text: bullet ? own.text : '•' })
    else update({ fmt, text: bullet || own.text === '' ? `%${level + 1}.` : own.text })
  }

  function apply(): void {
    editor.chain().focus().applyListLevels(kindOfLevels(levels), levels).run()
    onClose()
  }

  return (
    <div
      className="popover popover--wide"
      role="dialog"
      aria-label={t('document.lists.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <ListGallery levels={levels} onPick={setLevels} />

      <LevelFields level={level} own={own} onLevel={setLevel} onUpdate={update} onFormat={changeFormat} />

      <p className="popover__hint" aria-label={t('document.lists.preview')}>
        {t('document.lists.preview')}: <span className="list-preview">{preview.join('  ')}</span>
      </p>

      <NumberingActions editor={editor} ordered={ordered} onDone={onClose} />

      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={apply}>
          {t('document.common.apply')}
        </button>
      </div>
    </div>
  )
}

function ListGallery({
  levels,
  onPick,
}: {
  levels: readonly LevelDef[]
  onPick: (levels: LevelDef[]) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.lists.gallery')}</legend>
      <div className="list-gallery">
        {LIST_PRESETS.map((preset) => (
          <button
            key={preset.id}
            type="button"
            className="btn list-gallery__item"
            aria-pressed={JSON.stringify(preset.levels) === JSON.stringify(levels)}
            onClick={() => onPick(preset.levels.map((entry) => ({ ...entry })))}
          >
            {t(preset.labelKey as MessageKey)}
          </button>
        ))}
      </div>
    </fieldset>
  )
}

function LevelFields({
  level,
  own,
  onLevel,
  onUpdate,
  onFormat,
}: {
  level: number
  own: LevelDef
  onLevel: (level: number) => void
  onUpdate: (change: Partial<LevelDef>) => void
  onFormat: (fmt: string) => void
}): React.JSX.Element {
  const t = useT()
  const bullet = own.fmt === 'bullet'
  return (
    <div className="popover__row">
      <label className="popover__field popover__field--narrow">
        <span>{t('document.lists.level')}</span>
        <select value={level} onChange={(event) => onLevel(Number(event.target.value))}>
          {Array.from({ length: LIST_LEVELS }, (_, index) => (
            <option key={index} value={index}>
              {index + 1}
            </option>
          ))}
        </select>
      </label>

      <label className="popover__field">
        <span>{t('document.lists.numberFormat')}</span>
        <select value={own.fmt} onChange={(event) => onFormat(event.target.value)}>
          {NUMBER_FORMATS.map((fmt) => (
            <option key={fmt} value={fmt}>
              {t(FORMAT_LABELS[fmt])}
            </option>
          ))}
          {!(NUMBER_FORMATS as readonly string[]).includes(own.fmt) && (
            <option value={own.fmt}>{own.fmt}</option>
          )}
        </select>
      </label>

      {bullet ? (
        <label className="popover__field popover__field--narrow">
          <span>{t('document.lists.marker')}</span>
          <select value={own.text} onChange={(event) => onUpdate({ text: event.target.value })}>
            {[...new Set([...BULLET_MARKS, own.text])].map((mark) => (
              <option key={mark} value={mark}>
                {mark}
              </option>
            ))}
          </select>
        </label>
      ) : (
        <label className="popover__field">
          <span>{t('document.lists.levelText')}</span>
          <input
            type="text"
            value={own.text}
            disabled={own.fmt === 'none'}
            title={t('document.lists.levelTextHint')}
            onChange={(event) => onUpdate({ text: event.target.value })}
          />
        </label>
      )}

      {!bullet && (
        <label className="popover__field popover__field--narrow">
          <span>{t('document.lists.startAt')}</span>
          <input
            type="number"
            min={0}
            value={own.start}
            onChange={(event) =>
              onUpdate({ start: Math.max(0, Math.round(Number(event.target.value) || 0)) })
            }
          />
        </label>
      )}
    </div>
  )
}

/** They apply immediately, without Apply: they belong to the cursor's list, not the definition. */
function NumberingActions({
  editor,
  ordered,
  onDone,
}: {
  editor: Editor
  ordered: boolean
  onDone: () => void
}): React.JSX.Element {
  const t = useT()
  const [startValue, setStartValue] = useState('1')
  const start = Number(startValue)
  const validStart = Number.isInteger(start) && start >= 0 && start <= MAX_START_NUMBER
  const act = (run: () => boolean): void => {
    run()
    onDone()
  }
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.lists.numbering')}</legend>
      {ordered ? (
        <div className="popover__row">
          <button
            type="button"
            className="btn"
            onClick={() => act(() => editor.chain().focus().restartListNumbering(1).run())}
          >
            {t('document.lists.restart')}
          </button>
          <button
            type="button"
            className="btn"
            onClick={() => act(() => editor.chain().focus().continueListNumbering().run())}
          >
            {t('document.lists.continue')}
          </button>
          <label className="popover__field popover__field--narrow">
            <span>{t('document.lists.startAt')}</span>
            <input
              type="number"
              min={0}
              aria-label={t('document.lists.setStartTitle')}
              value={startValue}
              onChange={(event) => setStartValue(event.target.value)}
            />
          </label>
          <button
            type="button"
            className="btn"
            disabled={!validStart}
            onClick={() => act(() => editor.chain().focus().restartListNumbering(start).run())}
          >
            {t('document.lists.setStartApply')}
          </button>
        </div>
      ) : (
        <p className="popover__hint">{t('document.lists.notInList')}</p>
      )}
    </fieldset>
  )
}

/** Just the number, for those who came through the right click. */
export function ListStartDialog({
  editor,
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const [value, setValue] = useState('1')
  const start = Number(value)
  const valid = Number.isInteger(start) && start >= 0 && start <= MAX_START_NUMBER

  function apply(): void {
    if (!valid) return
    editor.chain().focus().restartListNumbering(start).run()
    onClose()
  }

  return (
    <div className="popover" role="dialog" aria-label={t('document.lists.setStartTitle')}>
      <label className="popover__field">
        <span>{t('document.lists.startAt')}</span>
        <input
          type="number"
          min={0}
          value={value}
          autoFocus
          onChange={(event) => setValue(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') apply()
            if (event.key === 'Escape') onClose()
          }}
        />
      </label>
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" disabled={!valid} onClick={apply}>
          {t('document.lists.setStartApply')}
        </button>
      </div>
    </div>
  )
}
