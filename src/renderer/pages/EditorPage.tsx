import { useWorkspace } from '../state/workspace.js'

/**
 * Editor da Fase 1.
 *
 * É deliberadamente uma área de texto simples: o objetivo desta fase é provar
 * o ciclo completo de arquivo (criar, abrir, salvar, fechar, avisar sobre
 * alterações), não editar com formatação. O editor rico entra na Fase 2 e
 * substitui só este componente — o restante da fase continua valendo.
 */
export function EditorPage(): React.JSX.Element {
  const content = useWorkspace((state) => state.content)
  const setContent = useWorkspace((state) => state.setContent)

  return (
    <div className="editor">
      <textarea
        className="editor__area"
        value={content}
        onChange={(event) => setContent(event.target.value)}
        spellCheck={false}
        autoFocus
        aria-label="Conteúdo do documento"
        placeholder="Digite aqui…"
      />
    </div>
  )
}
