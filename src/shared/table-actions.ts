/**
 * As ações de tabela, uma vez só.
 *
 * Três lugares precisam da mesma lista com os mesmos rótulos: o menu nativo
 * (processo main), o menu de contexto do documento (renderer) e o mapa que
 * executa o comando no editor. Três cópias divergiriam no primeiro rótulo
 * ajustado — e a divergência apareceria como "o menu diz uma coisa e o botão
 * direito diz outra".
 *
 * Aqui só moram **dados**: id, rótulo e se a ação exige o cursor dentro de uma
 * tabela. O que cada uma faz é o comando do TableKit, e mora no renderer; a
 * tecla, quando existe, mora na tabela de atalhos.
 */

export const TableAction = {
  /** Abre o diálogo que pergunta quantas linhas e colunas. */
  Insert: 'table-insert',
  RowBefore: 'table-row-before',
  RowAfter: 'table-row-after',
  DeleteRow: 'table-delete-row',
  ColumnBefore: 'table-column-before',
  ColumnAfter: 'table-column-after',
  DeleteColumn: 'table-delete-column',
  MergeCells: 'table-merge-cells',
  SplitCell: 'table-split-cell',
  /** `w:tblHeader`: a primeira linha se repete no alto de cada página. */
  ToggleHeaderRow: 'table-header-row',
  Delete: 'table-delete',
  /** Abre o diálogo de largura de coluna, bordas e sombreamento. */
  Properties: 'table-properties',
} as const

export type TableAction = (typeof TableAction)[keyof typeof TableAction]

export interface TableActionInfo {
  readonly id: TableAction
  readonly label: string
  /**
   * Se a ação só faz sentido com o cursor dentro de uma tabela.
   *
   * Decide o item apagado no menu e a ausência dele no menu de contexto: fora de
   * uma tabela, "mesclar células" não tem o que mesclar.
   */
  readonly needsTable: boolean
  /**
   * Grupo visual. O menu põe um separador quando o número muda — é o que separa
   * "inserir" de "excluir" e de "mesclar", como no Word.
   */
  readonly group: number
}

export const TABLE_ACTIONS: readonly TableActionInfo[] = [
  { id: TableAction.Insert, label: 'Inserir tabela…', needsTable: false, group: 0 },
  { id: TableAction.RowBefore, label: 'Inserir linha acima', needsTable: true, group: 1 },
  { id: TableAction.RowAfter, label: 'Inserir linha abaixo', needsTable: true, group: 1 },
  { id: TableAction.ColumnBefore, label: 'Inserir coluna à esquerda', needsTable: true, group: 1 },
  { id: TableAction.ColumnAfter, label: 'Inserir coluna à direita', needsTable: true, group: 1 },
  { id: TableAction.DeleteRow, label: 'Excluir linha', needsTable: true, group: 2 },
  { id: TableAction.DeleteColumn, label: 'Excluir coluna', needsTable: true, group: 2 },
  { id: TableAction.Delete, label: 'Excluir tabela', needsTable: true, group: 2 },
  { id: TableAction.MergeCells, label: 'Mesclar células', needsTable: true, group: 3 },
  { id: TableAction.SplitCell, label: 'Dividir célula', needsTable: true, group: 3 },
  { id: TableAction.ToggleHeaderRow, label: 'Linha de cabeçalho', needsTable: true, group: 4 },
  { id: TableAction.Properties, label: 'Propriedades da tabela…', needsTable: true, group: 4 },
]
