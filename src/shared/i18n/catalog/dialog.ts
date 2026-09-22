import type { Catalog } from '../message.js'

/**
 * Frases dos diálogos nativos do sistema operacional: abrir, salvar,
 * confirmação de descarte, aviso de texto simples e sobre.
 */
export const DIALOG = {
  // Filtros de arquivo
  'dialog.filter.allSupported': {
    pt: 'Todos os arquivos suportados',
    en: 'All supported files',
  },
  'dialog.filter.wordDocs': { pt: 'Documentos do Word', en: 'Word documents' },
  'dialog.filter.excelSheets': { pt: 'Planilhas do Excel', en: 'Excel spreadsheets' },
  'dialog.filter.documents': { pt: 'Documentos', en: 'Documents' },
  'dialog.filter.spreadsheets': { pt: 'Planilhas', en: 'Spreadsheets' },
  'dialog.filter.plainText': { pt: 'Texto simples', en: 'Plain text' },
  'dialog.filter.allFiles': { pt: 'Todos os arquivos', en: 'All files' },
  'dialog.filter.images': { pt: 'Imagens', en: 'Images' },
  'dialog.filter.document': { pt: 'Documento', en: 'Document' },
  'dialog.filter.wordDoc': { pt: 'Documento do Word', en: 'Word document' },
  'dialog.filter.spreadsheet': { pt: 'Planilha', en: 'Spreadsheet' },
  'dialog.filter.excelSheet': { pt: 'Planilha do Excel', en: 'Excel spreadsheet' },

  // Abrir e Salvar
  'dialog.open.title': { pt: 'Abrir arquivo', en: 'Open file' },
  'dialog.save.title': { pt: 'Salvar como', en: 'Save as' },
  'dialog.pdf.title': { pt: 'Exportar para PDF', en: 'Export to PDF' },
  'dialog.image.title': { pt: 'Inserir imagem', en: 'Insert image' },

  // Confirmar descarte de alterações
  'dialog.discard.title': { pt: 'Alterações não salvas', en: 'Unsaved changes' },
  'dialog.discard.message': {
    pt: 'Salvar as alterações em “{fileName}”?',
    en: 'Save changes to “{fileName}”?',
  },
  'dialog.discard.detail': {
    pt: 'Se não salvar, as alterações feitas desde a última gravação serão perdidas.',
    en: 'If you don’t save, changes made since the last save will be lost.',
  },
  'dialog.discard.save': { pt: 'Salvar', en: 'Save' },
  'dialog.discard.dontSave': { pt: 'Não salvar', en: 'Don’t save' },
  'dialog.discard.cancel': { pt: 'Cancelar', en: 'Cancel' },

  // Aviso de texto simples (.txt)
  'dialog.plainText.title': { pt: 'Formatação será perdida', en: 'Formatting will be lost' },
  'dialog.plainText.message': {
    pt: '“{fileName}” é um arquivo de texto simples.',
    en: '“{fileName}” is a plain text file.',
  },
  'dialog.plainText.detail': {
    pt: 'Texto simples não guarda negrito, títulos, listas, tabelas nem imagens. Salvar como documento preserva tudo.',
    en: 'Plain text does not keep bold, headings, lists, tables, or images. Saving as a document preserves everything.',
  },
  'dialog.plainText.saveAsDocument': {
    pt: 'Salvar como documento',
    en: 'Save as document',
  },
  'dialog.plainText.saveAsPlain': {
    pt: 'Salvar como texto simples',
    en: 'Save as plain text',
  },

  // Diálogo Sobre
  'dialog.about.title': { pt: 'Sobre o {app}', en: 'About {app}' },
  'dialog.about.detail': {
    pt: 'Versão {version}\n\nSuíte de documentos e planilhas, offline.\nEm desenvolvimento — Fase 7 de 8.',
    en: 'Version {version}\n\nDocument and spreadsheet suite, offline.\nIn development — Phase 7 of 8.',
  },
  'dialog.about.close': { pt: 'Fechar', en: 'Close' },
} satisfies Catalog
