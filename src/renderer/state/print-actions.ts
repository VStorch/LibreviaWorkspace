import type { PageSetup } from '@services/document/model.js'
import { buildPrintHtml } from '@services/document/print-html.js'
import { effectiveSections } from '@services/document/sections.js'
import { buildPagedBody, buildPagedCss } from '@services/document/print-pages.js'
import { styleSheetCss } from '@services/document/style-css.js'
import { SHEET_PRINT_CSS, buildSheetHtml } from '@services/spreadsheet/print-html.js'
import { t } from '../i18n.js'
import { currentPreferences } from './preferences.js'
import type { GetWorkspace, SetWorkspace, WorkspaceContext } from './context.js'
import type { WorkspaceState } from './types.js'

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
  /** O documento entrega o HTML do editor; a planilha é gerada do modelo, porque a grade só desenha as células visíveis. */
  function buildRequest(): PrintRequest | null {
    const state = get()
    const name = state.file?.name ?? t('shell.print.defaultDocumentName')

    const { workbook } = state
    if (workbook !== null) {
      const sheet = workbook.sheets[workbook.activeSheet]
      if (sheet === undefined) return null

      return {
        html: buildPrintHtml(buildSheetHtml(sheet, currentPreferences().language), name, SHEET_PRINT_CSS),
        page: state.page,
        // A grade é uma tabela contínua que o Chromium reparte.
        paged: false,
      }
    }

    const source = ctx.source()
    if (source === null) return null

    // O pedido leva o papel da primeira seção, que a impressora nativa oferece como padrão.
    const paged = source.readPages()
    const page = paged.pages[0]?.setup ?? effectiveSections(state.page, state.sections)[0]!

    return {
      html: buildPrintHtml(
        buildPagedBody(paged),
        // O `<title>` vira o Title do PDF; autor e assunto o `printToPDF` não grava.
        state.properties?.title?.trim() || name,
        styleSheetCss(state.styles) + buildPagedCss(paged.pages),
        false,
      ),
      page,
      paged: true,
    }
  }

  /** Devolver `false` em silêncio faria "Exportar para PDF" não dar em nada. */
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

      // Exportar não é salvar: nada aqui muda `file` nem `isDirty`.
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
