import type { Catalog } from '../message.js'

export const DIALOG = {
  'dialog.filter.allSupported': {
    pt: 'Todos os arquivos suportados',
    en: 'All supported files',
  },
  'dialog.filter.wordDocs': { pt: 'Documentos do Word', en: 'Word documents' },
  'dialog.filter.excelSheets': { pt: 'Planilhas do Excel', en: 'Excel spreadsheets' },
  'dialog.filter.documents': { pt: 'Documentos', en: 'Documents' },
  'dialog.filter.spreadsheets': { pt: 'Planilhas', en: 'Spreadsheets' },
  'dialog.filter.plainText': { pt: 'Texto simples', en: 'Plain text' },
  'dialog.filter.pdfDocs': { pt: 'Documentos PDF', en: 'PDF documents' },
  'dialog.filter.allFiles': { pt: 'Todos os arquivos', en: 'All files' },
  'dialog.filter.images': { pt: 'Imagens', en: 'Images' },
  'dialog.filter.document': { pt: 'Documento', en: 'Document' },
  'dialog.filter.wordDoc': { pt: 'Documento do Word', en: 'Word document' },
  'dialog.filter.wordTemplates': { pt: 'Modelos do Word', en: 'Word templates' },
  'dialog.filter.odtDocs': { pt: 'Documentos ODT', en: 'ODT documents' },
  'dialog.filter.wordTemplate': { pt: 'Modelo do Word (.dotx)', en: 'Word template (.dotx)' },
  'dialog.filter.odtTemplate': { pt: 'Modelo ODT (.ott)', en: 'ODT template (.ott)' },
  'dialog.filter.documentTemplates': { pt: 'Modelos de documento', en: 'Document templates' },
  'dialog.template.title': { pt: 'Procurar modelo', en: 'Browse for template' },
  'dialog.agent.title': { pt: 'Escolher o programa do agente', en: "Choose the agent's program" },
  'dialog.filter.programs': { pt: 'Programas', en: 'Programs' },
  'dialog.filter.spreadsheet': { pt: 'Planilha', en: 'Spreadsheet' },
  'dialog.filter.excelSheet': { pt: 'Planilha do Excel', en: 'Excel spreadsheet' },

  'dialog.open.title': { pt: 'Abrir arquivo', en: 'Open file' },
  'dialog.save.title': { pt: 'Salvar como', en: 'Save as' },
  'dialog.pdf.title': { pt: 'Exportar para PDF', en: 'Export to PDF' },
  'dialog.export.htmlTitle': { pt: 'Exportar como HTML', en: 'Export as HTML' },
  'dialog.export.markdownTitle': { pt: 'Exportar como Markdown', en: 'Export as Markdown' },
  'dialog.filter.html': { pt: 'Página da Web (HTML)', en: 'Web page (HTML)' },
  'dialog.filter.markdown': { pt: 'Markdown', en: 'Markdown' },
  'dialog.export.odtTitle': { pt: 'Exportar como ODT', en: 'Export as ODT' },
  'dialog.filter.odt': { pt: 'Documento ODT', en: 'ODT document' },
  // Written inside the exported file, in the UI language.
  'dialog.export.notes': { pt: 'Notas', en: 'Notes' },
  'dialog.export.backToText': { pt: 'Voltar ao texto', en: 'Back to text' },
  'dialog.image.title': { pt: 'Inserir imagem', en: 'Insert image' },

  'dialog.saveFormat.title': { pt: 'Formato dos arquivos novos', en: 'Format for new files' },
  'dialog.saveFormat.message': {
    pt: 'Em que formato salvar os arquivos novos?',
    en: 'Which format should new files be saved in?',
  },
  'dialog.saveFormat.detail': {
    pt: 'Office e Google: .docx e .xlsx, que Word, Excel, Google Docs e Planilhas abrem.\nLibreOffice: .odt para documentos; planilhas em .xlsx, que o Calc abre.\nLibrevia: .sdoc e .ssheet, que guardam tudo, mas só o Librevia abre.\n\nVocê pode mudar isso em Ferramentas → Opções.',
    en: 'Office and Google: .docx and .xlsx, which Word, Excel, Google Docs and Sheets open.\nLibreOffice: .odt for documents; spreadsheets as .xlsx, which Calc opens.\nLibrevia: .sdoc and .ssheet, which keep everything, but only Librevia opens.\n\nYou can change this in Tools → Options.',
  },
  'dialog.saveFormat.office': { pt: 'Office e Google', en: 'Office and Google' },
  'dialog.saveFormat.libreOffice': { pt: 'LibreOffice', en: 'LibreOffice' },
  'dialog.saveFormat.librevia': { pt: 'Librevia', en: 'Librevia' },

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

  'dialog.externalLink.title': { pt: 'Abrir link', en: 'Open link' },
  'dialog.externalLink.message': {
    pt: 'Abrir este endereço no navegador?',
    en: 'Open this address in the browser?',
  },
  'dialog.externalLink.open': { pt: 'Abrir', en: 'Open' },

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

  'dialog.about.title': { pt: 'Sobre o {app}', en: 'About {app}' },
  'dialog.about.detail': {
    pt: 'Versão {version}\n\nSuíte de documentos e planilhas, offline.',
    en: 'Version {version}\n\nDocument and spreadsheet suite, offline.',
  },
  'dialog.about.close': { pt: 'Fechar', en: 'Close' },
} satisfies Catalog
