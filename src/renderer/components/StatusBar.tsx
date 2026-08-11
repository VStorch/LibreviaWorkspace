import { isDirty, useWorkspace } from '../state/workspace.js'

export function StatusBar(): React.JSX.Element {
  const state = useWorkspace()
  const dirty = isDirty(state)
  const characters = state.content.length

  return (
    <footer className="statusbar">
      <span className="statusbar__file" title={state.file?.path ?? undefined}>
        {state.file?.path ?? 'Arquivo ainda não salvo'}
      </span>

      <span className="statusbar__spacer" />

      <span className="statusbar__metric">
        {characters.toLocaleString('pt-BR')} {characters === 1 ? 'caractere' : 'caracteres'}
      </span>

      {/* Indicador de alterações não salvas: repete o marcador do título da
          janela, para que o estado seja legível sem sair do conteúdo. */}
      <span className={dirty ? 'statusbar__state statusbar__state--dirty' : 'statusbar__state'}>
        {state.busy ? 'Trabalhando…' : dirty ? '• Não salvo' : 'Salvo'}
      </span>
    </footer>
  )
}
