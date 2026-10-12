import { DocumentKind, type LossInventory } from '@shared/types.js'
import { isOfficePackagePath } from '@services/file/formats.js'
import { propertiesOf } from '@services/file/properties.js'
import {
  PrivacyCategory,
  anonymizeAuthors,
  inspectOwnFile,
  inspectPdf,
  cleanOwnProperties,
  type PrivacyReport,
} from '@services/file/privacy.js'
import { t } from '../i18n.js'
import { interpret, openedNotice } from './file-actions.js'
import type { GetWorkspace, SetWorkspace, WorkspaceContext } from './context.js'
import type { OpenFile, WorkspaceState } from './types.js'

type PrivacyActions = Pick<WorkspaceState, 'inspectPersonalData' | 'removePersonalData'>

/** As Word names the authors it hides; main sends the sidecar the same name. */
export function anonymousName(): string {
  return t('revisions.unknownAuthor')
}

/** A DOCX or XLSX: what matters is in the package, which the sidecar reads from disk. */
function packagePathOf(file: OpenFile): string | null {
  return file.path !== null && isOfficePackagePath(file.path) ? file.path : null
}

/**
 * A package is inspected as saved: unsaved edits, or a document from a Word template that was never
 * saved, would not be in what the sidecar reads.
 */
export function inspectionNeedsSave(state: Pick<WorkspaceState, 'file' | 'isDirty'>): boolean {
  const { file } = state
  if (file === null || file.kind === DocumentKind.Pdf) return false
  const target = file.path ?? file.name
  return isOfficePackagePath(target) && (state.isDirty || file.path === null)
}

export function createPrivacyActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): PrivacyActions {
  /** The clean copy replaces the opened file, without the undo history, and waits to be saved. */
  async function cleanPackage(
    file: OpenFile,
    path: string,
    categories: readonly PrivacyCategory[],
  ): Promise<boolean> {
    const data = await ctx.call(() => window.api.privacy.remove({ path, categories: [...categories] }))
    if (data === null) return false

    const inventory: LossInventory | undefined = data.inventory
    ctx.show(interpret({ path, name: file.name, content: data.content, inventory }))
    set({ ...openedNotice(inventory), isDirty: true })
    return true
  }

  /** A file of the editor's own only has names and custom properties. */
  function cleanOwnFile(file: OpenFile, categories: readonly PrivacyCategory[]): boolean {
    const state = get()
    const properties = cleanOwnProperties(state.properties, categories)
    if (state.workbook === null && categories.includes(PrivacyCategory.Names)) {
      const model = anonymizeAuthors(ctx.currentModel(), anonymousName())
      ctx.show({ file, model: { ...model, properties }, workbook: null })
    }
    set({ properties, isDirty: true })
    return true
  }

  async function inspectOpenPdf(file: OpenFile): Promise<PrivacyReport | null> {
    if (file.path === null) return null
    const { path } = file
    const inspection = await ctx.call(() => window.api.pdf.properties({ path }))
    if (inspection === null) return null
    // The properties being edited, when the dialog already changed them.
    const properties = get().properties ?? propertiesOf(inspection.properties ?? {})
    return inspectPdf(properties, inspection.updates, anonymousName())
  }

  return {
    inspectPersonalData: async (): Promise<PrivacyReport | null> => {
      const state = get()
      if (state.file === null) return null
      if (state.file.kind === DocumentKind.Pdf) return inspectOpenPdf(state.file)

      const path = packagePathOf(state.file)
      if (path !== null) return ctx.call(() => window.api.privacy.inspect({ path }))

      const content = state.workbook === null ? ctx.currentModel() : null
      return inspectOwnFile(state.properties, content, anonymousName())
    },

    removePersonalData: async (categories) => {
      const { file, readOnly } = get()
      if (file === null || readOnly || categories.length === 0) return false

      const path = packagePathOf(file)
      return path === null ? cleanOwnFile(file, categories) : cleanPackage(file, path, categories)
    },
  }
}
