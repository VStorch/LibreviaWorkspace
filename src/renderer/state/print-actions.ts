import type { PageSetup } from '@services/document/model.js'
import { buildPrintHtml } from '@services/document/print-html.js'
import { buildPagedBody, buildPagedCss } from '@services/document/print-pages.js'
import { SHEET_PRINT_CSS, buildSheetHtml } from '@services/spreadsheet/print-html.js'
import type { GetWorkspace, SetWorkspace, WorkspaceContext } from './context.js'
import type { WorkspaceState } from './types.js'

/** O que o processo main precisa para pôr algo no papel. */
interface PrintRequest {
  readonly html: string
  readonly page: PageSetup
  readonly paged: boolean
}

type PrintActions = Pick<WorkspaceState, 'exportPdf' | 'print' | 'printPreview'>

export function createPrintActions(
  set: SetWorkspace,
  get: GetWorkspace,
  ctx: WorkspaceContext,
): PrintActions {
  /**
   * O que vai para o PDF e para a impressora.
   *
   * O documento entrega o HTML do próprio editor — o que sai no papel é
   * literalmente o que estava na tela. A planilha **não pode** fazer o mesmo: a
   * grade só desenha as células visíveis, então imprimir o DOM dela renderizaria
   * a janela, e não a planilha. Por isso a planilha é gerada a partir do modelo,
   * com as mesmas funções de formatação que a tela usa.
   */
  function buildRequest(): PrintRequest | null {
    const state = get()
    const name = state.file?.name ?? 'Documento'

    const { workbook } = state
    if (workbook !== null) {
      const sheet = workbook.sheets[workbook.activeSheet]
      if (sheet === undefined) return null

      return {
        html: buildPrintHtml(buildSheetHtml(sheet), name, SHEET_PRINT_CSS),
        page: state.page,
        // A grade continua sendo uma tabela contínua que o Chromium reparte:
        // não há folha para recortar antes de imprimir.
        paged: false,
      }
    }

    const source = ctx.source()
    if (source === null) return null

    return {
      html: buildPrintHtml(
        buildPagedBody(source.readPages(), state.page),
        name,
        buildPagedCss(state.page),
        false,
      ),
      page: state.page,
      // Diz ao processo main para não deixar o Chromium paginar nem desenhar
      // faixa: as folhas já vêm prontas no HTML.
      paged: true,
    }
  }

  /**
   * Não há o que imprimir.
   *
   * Devolver `false` em silêncio era o defeito: o usuário clicava em "Exportar
   * para PDF" e nada acontecia — nem papel, nem aviso, nem pista. Um menu que
   * não faz nada é pior que um menu ausente.
   */
  function refuse(): false {
    set({
      error: {
        code: 'INTERNAL',
        message: 'Não há nada aberto para imprimir. Abra ou crie um documento ou uma planilha primeiro.',
      },
    })
    return false
  }

  return {
    exportPdf: async () => {
      const request = buildRequest()
      if (request === null) return refuse()

      const data = await ctx.call(() =>
        window.api.print.exportPdf({ ...request, suggestedName: get().file?.name ?? 'documento' }),
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
