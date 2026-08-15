import { pageDimensionsMm } from '@services/document/model.js'
import { useWorkspace } from '../state/workspace.js'

export function StatusBar(): React.JSX.Element {
  const state = useWorkspace()
  const { width, height } = pageDimensionsMm(state.page)

  return (
    <footer className="statusbar">
      <span className="statusbar__file" title={state.file?.path ?? undefined}>
        {state.file?.path ?? 'Arquivo ainda não salvo'}
      </span>

      <span className="statusbar__spacer" />

      <span className="statusbar__metric">
        {state.page.size} {width} × {height} mm
      </span>
      {/* "≈" porque é medida na tela, não paginação de verdade: essa só
          acontece na exportação. Ver PageGuides.tsx. */}
      <span className="statusbar__metric" title="Estimativa — a paginação exata é a da exportação">
        ≈ {state.estimatedPages} {state.estimatedPages === 1 ? 'página' : 'páginas'}
      </span>
      <span className="statusbar__metric">
        {state.stats.words.toLocaleString('pt-BR')} {state.stats.words === 1 ? 'palavra' : 'palavras'}
      </span>
      <span className="statusbar__metric">
        {state.stats.characters.toLocaleString('pt-BR')}{' '}
        {state.stats.characters === 1 ? 'caractere' : 'caracteres'}
      </span>

      {/* Indicador de alterações não salvas: repete o marcador do título da
          janela, para que o estado seja legível sem sair do conteúdo. */}
      <span className={state.isDirty ? 'statusbar__state statusbar__state--dirty' : 'statusbar__state'}>
        {state.busy ? 'Trabalhando…' : state.isDirty ? '• Não salvo' : 'Salvo'}
      </span>
    </footer>
  )
}
