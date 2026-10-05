import type { PageSetup } from '@services/document/model.js'
import { buildPrintHtml } from '@services/document/print-html.js'
import { effectiveSections } from '@services/document/sections.js'
import { buildPagedBody, buildPagedCss } from '@services/document/print-pages.js'
import { styleSheetCss } from '@services/document/style-css.js'
import { SHEET_PRINT_CSS, buildSheetHtml } from '@services/spreadsheet/print-html.js'
import { t } from '../i18n.js'
import { currentPreferences } from './preferences.js'
import type { GetWorkspace, SetWorkspace, WorkspaceContext } from './context.js'
import type { DocumentSource, WorkspaceState } from './types.js'

interface PrintRequest {
  readonly html: string
  readonly page: PageSetup
  readonly paged: boolean
}

type PrintActions = Pick<WorkspaceState, 'exportPdf' | 'exportDocument' | 'print' | 'printPreview'>

export function createPrintActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): PrintActions {
  const buildRequest = (): PrintRequest | null => printRequestOf(get(), ctx.source())

  /** Returning `false` silently would make "Export to PDF" do nothing. */
  function refuse(): false {
    set({
      error: {
        code: 'INTERNAL',
        message: t('shell.print.nothingToPrint'),
      },
    })
    return false
  }

  return {
    exportPdf: async () => {
      const request = buildRequest()
      if (request === null) return refuse()

      const data = await ctx.call(() =>
        window.api.print.exportPdf({
          ...request,
          suggestedName: get().file?.name ?? t('shell.print.defaultDocumentName').toLowerCase(),
        }),
      )
      return data !== null && !data.canceled
    },

    exportDocument: async (format) => {
      const state = get()
      if (state.workbook !== null || ctx.source() === null) {
        set({ error: { code: 'INTERNAL', message: t('shell.export.documentOnly') } })
        return false
      }

      // Exporting is not saving: nothing here changes `file` or `isDirty`.
      const data = await ctx.call(() =>
        window.api.print.exportDocument({
          format,
          content: ctx.currentContent(),
          suggestedName: state.file?.name ?? t('shell.print.defaultDocumentName').toLowerCase(),
        }),
      )
      return data !== null && !data.canceled
    },

    print: async () => {
      const request = buildRequest()
      if (request === null) return refuse()

      const data = await ctx.call(() => window.api.print.dialog(request))
      return data !== null && data.printed
    },

    printPreview: async () => {
      const request = buildRequest()
      if (request === null) {
        refuse()
        return
      }

      await ctx.call(() => window.api.print.preview({ ...request, title: get().file?.name ?? 'Documento' }))
    },
  }
}

/**
 * A document hands over the editor HTML; a spreadsheet is built from the model, because the grid
 * only draws visible cells.
 */
function printRequestOf(state: WorkspaceState, source: DocumentSource | null): PrintRequest | null {
  const name = state.file?.name ?? t('shell.print.defaultDocumentName')

  const { workbook } = state
  if (workbook !== null) {
    const sheet = workbook.sheets[workbook.activeSheet]
    if (sheet === undefined) return null

    return {
      html: buildPrintHtml(buildSheetHtml(sheet, currentPreferences().language), name, SHEET_PRINT_CSS),
      page: state.page,
      // The grid is a continuous table that Chromium splits.
      paged: false,
    }
  }

  if (source === null) return null

  // The request carries the first section's paper, which the native printer offers as default.
  const paged = source.readPages()
  const page = paged.pages[0]?.setup ?? effectiveSections(state.page, state.sections)[0]!

  return {
    html: buildPrintHtml(
      buildPagedBody(paged),
      // `<title>` becomes the PDF Title; `printToPDF` does not write author or subject.
      state.properties?.title?.trim() || name,
      styleSheetCss(state.styles) + buildPagedCss(paged.pages),
      false,
    ),
    page,
    paged: true,
  }
}
