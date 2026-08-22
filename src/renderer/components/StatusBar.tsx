import { pageDimensionsMm } from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'

export function StatusBar(): React.JSX.Element {
  const state = useWorkspace()
  const { width, height } = pageDimensionsMm(state.page)

  // Planilha não tem página, palavra nem caractere: mostrar isso ali seria
  // informação falsa ocupando o lugar da verdadeira.
  const sheet = state.workbook?.sheets[state.workbook.activeSheet]

  return (
    <footer className="statusbar">
      <span className="statusbar__file" title={state.file?.path ?? undefined}>
        {state.file?.path ?? 'Arquivo ainda não salvo'}
      </span>

      <span className="statusbar__spacer" />

      {sheet === undefined ? (
        <>
          <span className="statusbar__metric">
            {state.page.size} {width} × {height} mm
          </span>
          {/* Sem "≈": a tela agora pagina de verdade, e o número é o mesmo
              que a pessoa vê nas folhas. Prometer aproximação quando a conta
              está certa ensina a desconfiar de um número bom. */}
          <span className="statusbar__metric">
            {state.estimatedPages} {state.estimatedPages === 1 ? 'página' : 'páginas'}
          </span>
          <span className="statusbar__metric">
            {state.stats.words.toLocaleString('pt-BR')} {state.stats.words === 1 ? 'palavra' : 'palavras'}
          </span>
          <span className="statusbar__metric">
            {state.stats.characters.toLocaleString('pt-BR')}{' '}
            {state.stats.characters === 1 ? 'caractere' : 'caracteres'}
          </span>
        </>
      ) : (
        <>
          <span className="statusbar__metric">{sheet.name}</span>
          <span className="statusbar__metric">
            {Object.keys(sheet.cells).length.toLocaleString('pt-BR')}{' '}
            {Object.keys(sheet.cells).length === 1 ? 'célula preenchida' : 'células preenchidas'}
          </span>
        </>
      )}

      {/* Indicador de alterações não salvas: repete o marcador do título da
          janela, para que o estado seja legível sem sair do conteúdo. */}
      <span className={state.isDirty ? 'statusbar__state statusbar__state--dirty' : 'statusbar__state'}>
        {state.busy ? 'Trabalhando…' : state.isDirty ? '• Não salvo' : 'Salvo'}
      </span>
    </footer>
  )
}
