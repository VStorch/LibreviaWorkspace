import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { isValidParagraphDraft } from '@services/document/paragraph-format.js'
import {
  createStyle,
  styleDraftOf,
  styleFromBlock,
  styleWithDraft,
  withIdentity,
  type StyleDraft,
} from '@services/document/style-editing.js'
import { StyleType, listedStyles, styleLabelOf, type StyleSheet } from '@services/document/styles.js'
import { useLanguage, useT } from '../i18n.js'
import { blockAttrsOf } from './extensions/paragraph-commands.js'
import { ParagraphIndentFields } from './toolbar/ParagraphIndentFields.js'
import { ParagraphSpacingFields } from './toolbar/ParagraphSpacingFields.js'
import { ParagraphLineFields } from './toolbar/ParagraphLineFields.js'

export type StyleDialogMode =
  | { readonly kind: 'modify'; readonly id: string }
  | { readonly kind: 'create'; readonly basedOn: string | null }

/**
 * Shows what the style is **worth** and saves only what changed (`styleWithDraft`). The paragraph
 * fields are the paragraph dialog's. "Update to match selection" only fills the form; nothing
 * changes until OK.
 */
export function StyleDialog({
  editor,
  sheet,
  mode,
  onDone,
  onCancel,
}: {
  readonly editor: Editor
  readonly sheet: StyleSheet
  readonly mode: StyleDialogMode
  readonly onDone: (sheet: StyleSheet, id: string) => void
  readonly onCancel: () => void
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const editing = mode.kind === 'modify' ? sheet.styles[mode.id] : undefined

  const form = useStyleForm(sheet, mode)
  const { name, type, basedOn, next, draft, setDraft } = form

  const paragraphs = listedStyles(sheet, language).filter(
    (style) => style.type === StyleType.Paragraph && style.id !== editing?.id,
  )
  const nameEditable = mode.kind === 'create' || editing?.custom === true
  const valid = name.trim() !== '' && isValidParagraphDraft(draft.paragraph)

  function fromSelection(): void {
    const id = mode.kind === 'modify' ? mode.id : basedOn
    const node = editor.state.selection.$from.parent
    setDraft(styleFromBlock(sheet, id, blockAttrsOf(editor, node)))
  }

  function submit(): void {
    if (!valid) return
    const identity = {
      name,
      basedOn: basedOn === '' ? undefined : basedOn,
      next: type === StyleType.Paragraph && next !== '' ? next : undefined,
    }

    if (mode.kind === 'modify') {
      const renamed = withIdentity(sheet, mode.id, identity)
      onDone(styleWithDraft(renamed, mode.id, draft), mode.id)
      return
    }

    const created = createStyle(sheet, { ...identity, type })
    onDone(styleWithDraft(created.sheet, created.id, draft), created.id)
  }

  return (
    <div
      className="style-dialog"
      role="group"
      aria-label={
        mode.kind === 'modify' ? t('document.styles.modifyTitle') : t('document.styles.createTitle')
      }
      onKeyDown={(event) => {
        if (event.key === 'Enter' && (event.target as HTMLElement).tagName === 'INPUT') submit()
      }}
    >
      <IdentityFields form={form} mode={mode} paragraphs={paragraphs} nameEditable={nameEditable} />

      <FontFields draft={draft} setDraft={setDraft} />

      {type === StyleType.Paragraph && <ParagraphFields draft={draft} setDraft={setDraft} />}

      {name.trim() === '' && <p className="popover__error">{t('document.styles.nameMissing')}</p>}

      <div className="popover__actions">
        <button type="button" className="btn" onClick={fromSelection}>
          {t('document.styles.fromSelection')}
        </button>
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onCancel}>
          {t('document.common.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={submit} disabled={!valid}>
          {t('document.styles.ok')}
        </button>
      </div>
    </div>
  )
}

interface StyleForm {
  readonly name: string
  readonly setName: (name: string) => void
  readonly type: StyleType
  readonly setType: (type: StyleType) => void
  readonly basedOn: string
  readonly setBasedOn: (id: string) => void
  readonly next: string
  readonly setNext: (id: string) => void
  readonly draft: StyleDraft
  readonly setDraft: (draft: StyleDraft) => void
}

function useStyleForm(sheet: StyleSheet, mode: StyleDialogMode): StyleForm {
  const editing = mode.kind === 'modify' ? sheet.styles[mode.id] : undefined
  const fallback = sheet.defaults.paragraphStyleId
  const [name, setName] = useState(editing?.name ?? '')
  const [type, setType] = useState<StyleType>(editing?.type ?? StyleType.Paragraph)
  const [basedOn, setBasedOn] = useState<string>(
    editing?.basedOn ?? (mode.kind === 'create' ? (mode.basedOn ?? fallback ?? '') : ''),
  )
  const [next, setNext] = useState<string>(editing?.next ?? '')
  const [draft, setDraft] = useState<StyleDraft>(() => {
    const source = mode.kind === 'modify' ? mode.id : (mode.basedOn ?? fallback)
    return source === null ? styleDraftOf(sheet, '') : styleDraftOf(sheet, source)
  })

  return { name, setName, type, setType, basedOn, setBasedOn, next, setNext, draft, setDraft }
}

function IdentityFields({
  form,
  mode,
  paragraphs,
  nameEditable,
}: {
  form: StyleForm
  mode: StyleDialogMode
  paragraphs: ReturnType<typeof listedStyles>
  nameEditable: boolean
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const { name, setName, type, setType, basedOn, setBasedOn, next, setNext } = form
  return (
    <>
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('document.styles.name')}</span>
          <input
            aria-label={t('document.styles.name')}
            value={name}
            disabled={!nameEditable}
            autoFocus
            onChange={(event) => setName(event.target.value)}
          />
        </label>

        {mode.kind === 'create' && (
          <label className="popover__field">
            <span>{t('document.styles.type')}</span>
            <select
              aria-label={t('document.styles.type')}
              value={type}
              onChange={(event) => setType(event.target.value as StyleType)}
            >
              <option value={StyleType.Paragraph}>{t('document.styles.paragraphFilter')}</option>
              <option value={StyleType.Character}>{t('document.styles.characterFilter')}</option>
            </select>
          </label>
        )}
      </div>
      {!nameEditable && <p className="popover__hint">{t('document.styles.nameBuiltin')}</p>}

      <div className="popover__row">
        <label className="popover__field">
          <span>{t('document.styles.basedOnLabel')}</span>
          <select
            aria-label={t('document.styles.basedOnLabel')}
            value={basedOn}
            onChange={(event) => setBasedOn(event.target.value)}
          >
            <option value="">{t('document.styles.none')}</option>
            {paragraphs.map((style) => (
              <option key={style.id} value={style.id}>
                {styleLabelOf(style, language)}
              </option>
            ))}
          </select>
        </label>

        {type === StyleType.Paragraph && (
          <label className="popover__field">
            <span>{t('document.styles.nextLabel')}</span>
            <select
              aria-label={t('document.styles.nextLabel')}
              value={next}
              onChange={(event) => setNext(event.target.value)}
            >
              <option value="">{t('document.styles.none')}</option>
              {paragraphs.map((style) => (
                <option key={style.id} value={style.id}>
                  {styleLabelOf(style, language)}
                </option>
              ))}
            </select>
          </label>
        )}
      </div>
    </>
  )
}

function FontFields({
  draft,
  setDraft,
}: {
  draft: StyleDraft
  setDraft: (draft: StyleDraft) => void
}): React.JSX.Element {
  const t = useT()
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('document.styles.fontLegend')}</legend>
      <div className="popover__row">
        <label className="popover__field">
          <span>{t('document.styles.fontFamily')}</span>
          <input
            aria-label={t('document.styles.fontFamily')}
            value={draft.fontFamily}
            onChange={(event) => setDraft({ ...draft, fontFamily: event.target.value })}
          />
        </label>
        <label className="popover__field popover__field--narrow">
          <span>{t('document.styles.fontSize')}</span>
          <input
            type="number"
            min={1}
            max={400}
            step={0.5}
            aria-label={t('document.styles.fontSize')}
            value={draft.fontSize ?? ''}
            onChange={(event) =>
              setDraft({
                ...draft,
                fontSize: event.target.value === '' ? null : Number(event.target.value),
              })
            }
          />
        </label>
      </div>
      <div className="popover__row">
        {(['bold', 'italic', 'underline'] as const).map((key) => (
          <label key={key} className="popover__check">
            <input
              type="checkbox"
              checked={draft[key]}
              onChange={(event) => setDraft({ ...draft, [key]: event.target.checked })}
            />
            <span>{t(`document.characterFormat.${key}`)}</span>
          </label>
        ))}
      </div>
    </fieldset>
  )
}

function ParagraphFields({
  draft,
  setDraft,
}: {
  draft: StyleDraft
  setDraft: (draft: StyleDraft) => void
}): React.JSX.Element {
  const t = useT()
  const paragraph = draft.paragraph
  const changeParagraph = <K extends keyof StyleDraft['paragraph']>(
    key: K,
    value: StyleDraft['paragraph'][K],
  ) => setDraft({ ...draft, paragraph: { ...draft.paragraph, [key]: value } })
  return (
    <>
      <ParagraphLineFields
        draft={paragraph}
        onChange={changeParagraph}
        onLineSpacing={(spacing) => setDraft({ ...draft, paragraph: { ...paragraph, ...spacing } })}
      />

      <ParagraphSpacingFields draft={paragraph} onChange={changeParagraph} />
      <ParagraphIndentFields draft={paragraph} onChange={changeParagraph} />

      <label className="popover__check">
        <input
          type="checkbox"
          checked={paragraph.keepNext}
          onChange={(event) => changeParagraph('keepNext', event.target.checked)}
        />
        <span>{t('document.paragraph.keepWithNext')}</span>
      </label>
    </>
  )
}
