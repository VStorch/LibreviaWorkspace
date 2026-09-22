import type { Catalog } from '../message.js'

/**
 * Frases da planilha: barra de ferramentas, abas, menu de contexto e mensagens
 * de leitura e gravação do formato.
 */
export const SPREADSHEET = {
  // Barra de fórmulas
  'spreadsheet.formulaBar.selectedCell': { pt: 'Célula selecionada', en: 'Selected cell' },
  'spreadsheet.formulaBar.inputLabel': {
    pt: 'Fórmula ou conteúdo da célula',
    en: 'Formula or cell content',
  },

  // Menu de contexto
  'spreadsheet.contextMenu.label': { pt: 'Ações da planilha', en: 'Sheet actions' },
  'spreadsheet.rows.insertAbove': {
    pt: { one: 'Inserir linha acima', other: 'Inserir {count} linhas acima' },
    en: { one: 'Insert row above', other: 'Insert {count} rows above' },
  },
  'spreadsheet.rows.insertBelow': {
    pt: { one: 'Inserir linha abaixo', other: 'Inserir {count} linhas abaixo' },
    en: { one: 'Insert row below', other: 'Insert {count} rows below' },
  },
  'spreadsheet.rows.delete': {
    pt: { one: 'Excluir linha', other: 'Excluir {count} linhas' },
    en: { one: 'Delete row', other: 'Delete {count} rows' },
  },
  'spreadsheet.columns.insertLeft': {
    pt: { one: 'Inserir coluna à esquerda', other: 'Inserir {count} colunas à esquerda' },
    en: { one: 'Insert column to the left', other: 'Insert {count} columns to the left' },
  },
  'spreadsheet.columns.insertRight': {
    pt: { one: 'Inserir coluna à direita', other: 'Inserir {count} colunas à direita' },
    en: { one: 'Insert column to the right', other: 'Insert {count} columns to the right' },
  },
  'spreadsheet.columns.delete': {
    pt: { one: 'Excluir coluna', other: 'Excluir {count} colunas' },
    en: { one: 'Delete column', other: 'Delete {count} columns' },
  },
  'spreadsheet.contextMenu.clearContents': { pt: 'Limpar conteúdo', en: 'Clear contents' },

  // Abas
  'spreadsheet.sheetTabs.doubleClickToRename': {
    pt: 'Clique duplo para renomear',
    en: 'Double-click to rename',
  },
  'spreadsheet.sheetTabs.deleteSheet': { pt: 'Excluir {name}', en: 'Delete {name}' },
  'spreadsheet.sheetTabs.addSheet': { pt: 'Nova planilha na pasta', en: 'New sheet in workbook' },

  // Barra de ferramentas
  'spreadsheet.toolbar.label': { pt: 'Formatação da planilha', en: 'Spreadsheet formatting' },
  'spreadsheet.toolbar.textFormat': { pt: 'Formatação do texto', en: 'Text formatting' },
  'spreadsheet.toolbar.bold': { pt: 'Negrito', en: 'Bold' },
  'spreadsheet.toolbar.italic': { pt: 'Itálico', en: 'Italic' },
  'spreadsheet.toolbar.underline': { pt: 'Sublinhado', en: 'Underline' },
  'spreadsheet.toolbar.textColor': { pt: 'Cor do texto', en: 'Text color' },
  'spreadsheet.toolbar.fillColor': { pt: 'Cor de fundo', en: 'Fill color' },
  'spreadsheet.toolbar.alignment': { pt: 'Alinhamento', en: 'Alignment' },
  'spreadsheet.toolbar.alignLeft': { pt: 'Alinhar à esquerda', en: 'Align left' },
  'spreadsheet.toolbar.alignCenter': { pt: 'Centralizar', en: 'Center' },
  'spreadsheet.toolbar.alignRight': { pt: 'Alinhar à direita', en: 'Align right' },
  'spreadsheet.toolbar.numberGroup': { pt: 'Número', en: 'Number' },
  'spreadsheet.toolbar.numberFormat': { pt: 'Formato do número', en: 'Number format' },
  'spreadsheet.toolbar.decreaseDecimals': {
    pt: 'Menos casas decimais',
    en: 'Decrease decimal places',
  },
  'spreadsheet.toolbar.increaseDecimals': {
    pt: 'Mais casas decimais',
    en: 'Increase decimal places',
  },
  'spreadsheet.toolbar.borders': { pt: 'Bordas', en: 'Borders' },
  'spreadsheet.toolbar.allBorders': { pt: 'Bordas em volta', en: 'All borders' },
  'spreadsheet.toolbar.noBorders': { pt: 'Sem bordas', en: 'No borders' },
  'spreadsheet.toolbar.panes': { pt: 'Painéis', en: 'Panes' },
  'spreadsheet.toolbar.freeze': { pt: 'Congelar até a seleção', en: 'Freeze panes to selection' },
  'spreadsheet.toolbar.unfreeze': { pt: 'Descongelar', en: 'Unfreeze panes' },

  // Formatos de número
  'spreadsheet.format.general': { pt: 'Geral', en: 'General' },
  'spreadsheet.format.number': { pt: 'Número', en: 'Number' },
  'spreadsheet.format.currency': { pt: 'Moeda', en: 'Currency' },
  'spreadsheet.format.percent': { pt: 'Percentual', en: 'Percentage' },
  'spreadsheet.format.date': { pt: 'Data', en: 'Date' },
  'spreadsheet.format.text': { pt: 'Texto', en: 'Text' },

  // Impressão
  'spreadsheet.print.emptyTab': { pt: 'Esta aba está vazia.', en: 'This sheet is empty.' },

  // Serialização / erros
  'spreadsheet.error.corrupt': {
    pt: 'Esta planilha não pôde ser lida: o conteúdo está corrompido ou não é uma planilha válida.',
    en: 'This spreadsheet could not be read: the content is corrupt or not a valid spreadsheet.',
  },
  'spreadsheet.error.invalid': {
    pt: 'Este arquivo não é uma planilha válida deste aplicativo.',
    en: 'This file is not a valid spreadsheet for this application.',
  },
  'spreadsheet.error.newerVersion': {
    pt: 'Esta planilha foi criada por uma versão mais recente do aplicativo. Atualize para abri-la.',
    en: 'This spreadsheet was created by a newer version of the application. Update to open it.',
  },
} satisfies Catalog
