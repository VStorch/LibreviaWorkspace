import type { Catalog } from '../message.js'

/**
 * Frases da casca do aplicativo: faixas de aviso, barra de status, tela
 * inicial, títulos de janela e mensagens de estado global.
 */
export const SHELL = {
  // Arquivo / nomes padrão
  'shell.file.untitled': { pt: 'Sem título', en: 'Untitled' },
  'shell.file.untitledDocument': { pt: 'Documento sem título', en: 'Untitled document' },
  'shell.file.untitledSpreadsheet': { pt: 'Planilha sem título', en: 'Untitled spreadsheet' },

  // Faixa de erro
  'shell.banner.dismiss': { pt: 'Dispensar aviso', en: 'Dismiss notice' },

  // Faixa de inventário (perda cirúrgica)
  'shell.banner.savedLoss': {
    pt: 'Nesta gravação, isto não chegou ao arquivo:',
    en: 'In this save, this did not make it into the file:',
  },
  'shell.banner.willBeLost': {
    pt: 'Ao salvar, isto será perdido:',
    en: 'When saving, this will be lost:',
  },
  'shell.banner.stillInFileAnd': {
    pt: 'E isto continua no arquivo, mas não aparece por inteiro aqui:',
    en: 'And this remains in the file, but does not appear completely here:',
  },
  'shell.banner.stillInFileDoc': {
    pt: 'Este documento tem recursos que continuam no arquivo, mas não aparecem por inteiro aqui:',
    en: 'This document has features that remain in the file, but do not appear completely here:',
  },

  // Faixa de somente leitura
  'shell.banner.readOnly': {
    pt: 'Aberto somente para leitura.',
    en: 'Opened as read-only.',
  },
  'shell.banner.readOnlyReasons': {
    pt: 'Este arquivo tem {reasons} — que o editor não reproduz por inteiro. Tudo isso volta intacto ao salvar, menos o que estiver no trecho que você editar.',
    en: 'This file contains {reasons} — which the editor does not reproduce completely. All of this remains intact upon saving, except for what is in the section you edit.',
  },
  'shell.banner.readOnlyGeneric': {
    pt: 'Este arquivo tem recursos que o editor não reproduz por inteiro e que podem se perder ao editar.',
    en: 'This file contains features that the editor does not reproduce completely and may be lost when editing.',
  },
  'shell.banner.editAnyway': {
    pt: 'Editar mesmo assim',
    en: 'Edit anyway',
  },

  // Faixa de recuperação
  'shell.recovery.title': {
    pt: 'O aplicativo fechou com trabalho não salvo.',
    en: 'The application closed with unsaved work.',
  },
  'shell.recovery.detail': {
    pt: 'Há {what} de “{name}”, guardado {time}. Recuperar traz o conteúdo de volta para a tela sem gravar em nada — você decide onde salvar.',
    en: 'There is {what} from “{name}”, saved {time}. Recovering brings the content back to the screen without saving anywhere — you decide where to save.',
  },
  'shell.recovery.document': { pt: 'um documento', en: 'a document' },
  'shell.recovery.spreadsheet': { pt: 'uma planilha', en: 'a spreadsheet' },
  'shell.recovery.justNow': { pt: 'agora há pouco', en: 'just now' },
  'shell.recovery.minutesAgo': {
    pt: { one: 'há {count} minuto', other: 'há {count} minutos' },
    en: { one: '{count} minute ago', other: '{count} minutes ago' },
  },
  'shell.recovery.hoursAgo': {
    pt: { one: 'há {count} hora', other: 'há {count} horas' },
    en: { one: '{count} hour ago', other: '{count} hours ago' },
  },
  'shell.recovery.onDate': { pt: 'em {date}', en: 'on {date}' },
  'shell.recovery.recover': { pt: 'Recuperar', en: 'Recover' },
  'shell.recovery.discard': { pt: 'Descartar', en: 'Discard' },

  // Barra de status
  'shell.statusBar.zoomIn': { pt: 'Ampliar', en: 'Zoom in' },
  'shell.statusBar.zoomOut': { pt: 'Reduzir', en: 'Zoom out' },
  'shell.statusBar.zoomLevel': { pt: 'Zoom', en: 'Zoom' },
  'shell.statusBar.zoomFitWidth': { pt: 'Ajustar à largura', en: 'Fit to width' },
  'shell.statusBar.notSavedYet': { pt: 'Arquivo ainda não salvo', en: 'File not saved yet' },
  'shell.statusBar.pages': {
    pt: { one: '{count} página', other: '{count} páginas' },
    en: { one: '{count} page', other: '{count} pages' },
  },
  'shell.statusBar.words': {
    pt: { one: '{count} palavra', other: '{count} palavras' },
    en: { one: '{count} word', other: '{count} words' },
  },
  'shell.statusBar.characters': {
    pt: { one: '{count} caractere', other: '{count} caracteres' },
    en: { one: '{count} character', other: '{count} characters' },
  },
  'shell.statusBar.filledCells': {
    pt: { one: '{count} célula preenchida', other: '{count} células preenchidas' },
    en: { one: '{count} filled cell', other: '{count} filled cells' },
  },
  'shell.statusBar.working': { pt: 'Trabalhando…', en: 'Working…' },
  'shell.statusBar.unsaved': { pt: '• Não salvo', en: '• Unsaved' },
  'shell.statusBar.saved': { pt: 'Salvo', en: 'Saved' },

  // Tela inicial (HomePage)
  'shell.home.subtitle': {
    pt: 'Documentos e planilhas, sem depender de nuvem.',
    en: 'Documents and spreadsheets, without depending on the cloud.',
  },
  'shell.home.newDocument': { pt: 'Novo documento', en: 'New document' },
  'shell.home.newSpreadsheet': { pt: 'Nova planilha', en: 'New spreadsheet' },
  'shell.home.openFile': { pt: 'Abrir arquivo', en: 'Open file' },
  'shell.home.recentFiles': { pt: 'Arquivos recentes', en: 'Recent files' },
  'shell.home.clear': { pt: 'Limpar', en: 'Clear' },
  'shell.home.emptyRecents': { pt: 'Nenhum arquivo aberto ainda.', en: 'No files opened yet.' },

  // Erros e avisos de estado
  'shell.error.unexpected': {
    pt: 'Ocorreu um erro inesperado. A operação não foi concluída.',
    en: 'An unexpected error occurred. The operation could not be completed.',
  },
  'shell.error.communicationFailed': {
    pt: 'A comunicação com o aplicativo falhou. Tente novamente.',
    en: 'Communication with the application failed. Please try again.',
  },
  'shell.error.rootNotFound': {
    pt: 'Elemento #root não encontrado.',
    en: '#root element not found.',
  },
  'shell.draft.autosaveBroken': {
    pt: 'A gravação automática de segurança parou de funcionar. Seu arquivo não foi alterado, mas salve o trabalho manualmente: uma queda agora custaria o que não foi salvo.',
    en: 'Automatic backup saving stopped working. Your file was not modified, but save your work manually: a crash now would cost what was not saved.',
  },
  'shell.print.nothingToPrint': {
    pt: 'Não há nada aberto para imprimir. Abra ou crie um documento ou uma planilha primeiro.',
    en: 'Nothing is open to print. Open or create a document or spreadsheet first.',
  },
  'shell.print.defaultDocumentName': {
    pt: 'Documento',
    en: 'Document',
  },
} satisfies Catalog
