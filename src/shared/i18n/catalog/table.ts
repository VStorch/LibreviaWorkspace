import type { Catalog } from '../message.js'

/**
 * As ações de tabela.
 *
 * Área própria porque a lista em `shared/table-actions.ts` é **dado**, e é lida
 * por três consumidores — o menu nativo, o menu de contexto e o mapa de
 * comandos. Enquanto o rótulo era uma string dentro do dado, traduzi-lo exigiria
 * traduzir o dado; com a chave no lugar do rótulo, o dado continua sendo dado e
 * cada consumidor resolve a frase no idioma que ele conhece.
 *
 * É o padrão para toda tabela de dados com rótulo neste repositório.
 */
export const TABLE = {
  'table.insert': { pt: 'Inserir tabela…', en: 'Insert table…' },
  'table.rowBefore': { pt: 'Inserir linha acima', en: 'Insert row above' },
  'table.rowAfter': { pt: 'Inserir linha abaixo', en: 'Insert row below' },
  'table.columnBefore': { pt: 'Inserir coluna à esquerda', en: 'Insert column to the left' },
  'table.columnAfter': { pt: 'Inserir coluna à direita', en: 'Insert column to the right' },
  'table.deleteRow': { pt: 'Excluir linha', en: 'Delete row' },
  'table.deleteColumn': { pt: 'Excluir coluna', en: 'Delete column' },
  'table.delete': { pt: 'Excluir tabela', en: 'Delete table' },
  'table.mergeCells': { pt: 'Mesclar células', en: 'Merge cells' },
  'table.splitCell': { pt: 'Dividir célula', en: 'Split cell' },
  'table.headerRow': { pt: 'Linha de cabeçalho', en: 'Header row' },
  'table.properties': { pt: 'Propriedades da tabela…', en: 'Table properties…' },
} satisfies Catalog
