/**
 * Canal interno entre o menu nativo e o editor.
 *
 * O menu vive no processo main e chega ao renderer pelo `App`, que não tem
 * referência ao editor. Um emissor mínimo resolve isso sem colocar o objeto do
 * editor — que não é serializável — dentro do store.
 *
 * Os identificadores são **os mesmos** dos comandos de menu correspondentes
 * (`MenuCommand`), e de propósito: o `App` reconhece um comando de editor pelo
 * nome e o repassa, em vez de manter uma tradução de um para um que cresce um
 * caso a cada recurso. Foram doze de uma vez com as tabelas.
 */

import { TableAction } from '@shared/table-actions.js'
import type { MenuCommand } from '@shared/types.js'

/**
 * O `satisfies` é a outra ponta do repasse por nome: um valor que não exista em
 * `MenuCommand` — renomeado de um lado só, digitado errado — não compila. Sem
 * ele o comando chegava ao `App`, não era reconhecido como do editor e virava
 * nada, calado.
 */
export const EditorCommand = {
  FindReplace: 'find-replace',
  PageSetup: 'page-setup',
  InsertPageBreak: 'insert-page-break',
  ParagraphSetup: 'paragraph-setup',
  PasteWithoutFormat: 'paste-without-format',
  WordCount: 'word-count',
  SpecialCharacter: 'special-character',
  /** Propriedades da imagem selecionada: texto alternativo e alinhamento. */
  ImageProperties: 'image-properties',
  /** Marcadores: adicionar, ir para e excluir. */
  InsertBookmark: 'insert-bookmark',
  InsertTableOfContents: 'insert-table-of-contents',
  UpdateTableOfContents: 'update-table-of-contents',
  UpdateFields: 'update-fields',
  ...TableAction,
} as const satisfies Record<string, MenuCommand>

export type EditorCommand = (typeof EditorCommand)[keyof typeof EditorCommand]

const KNOWN = new Set<string>(Object.values(EditorCommand))

/**
 * O comando de menu que, na verdade, é do editor — ou `null`.
 *
 * É o que permite ao `App` repassar sem conhecer: tudo o que precisa de seleção,
 * cursor ou diálogo do documento é do editor, e o `App` só sabe de arquivos.
 */
export function asEditorCommand(command: string): EditorCommand | null {
  return KNOWN.has(command) ? (command as EditorCommand) : null
}

/**
 * Os comandos que só leem o documento, e por isso valem com ele travado.
 *
 * A lista é a das exceções, e não a das edições, de propósito: um comando novo
 * nasce bloqueado no somente leitura até alguém decidir que ele não edita. O
 * contrário — esquecer de acrescentar uma edição à lista — era o furo por onde
 * os comandos de tabela e a quebra de página passavam.
 */
const READS_ONLY: ReadonlySet<EditorCommand> = new Set<EditorCommand>([
  EditorCommand.FindReplace,
  EditorCommand.WordCount,
  // O diálogo abre para "Ir para"; adicionar e excluir se apagam lá dentro.
  EditorCommand.InsertBookmark,
])

/** O comando pode rodar num documento aberto em somente leitura? */
export function runsWhileLocked(command: EditorCommand): boolean {
  return READS_ONLY.has(command)
}

const listeners = new Set<(command: EditorCommand) => void>()

export function onEditorCommand(listener: (command: EditorCommand) => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

export function emitEditorCommand(command: EditorCommand): void {
  for (const listener of listeners) listener(command)
}
