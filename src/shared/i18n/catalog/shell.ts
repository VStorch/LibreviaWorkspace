import type { Catalog } from '../message.js'

export const SHELL = {
  'shell.file.untitled': { pt: 'Sem título', en: 'Untitled' },
  'shell.file.untitledDocument': { pt: 'Documento sem título', en: 'Untitled document' },
  'shell.file.untitledSpreadsheet': { pt: 'Planilha sem título', en: 'Untitled spreadsheet' },

  'shell.banner.dismiss': { pt: 'Dispensar aviso', en: 'Dismiss notice' },

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

  'shell.statusBar.zoomIn': { pt: 'Ampliar', en: 'Zoom in' },
  'shell.statusBar.zoomOut': { pt: 'Reduzir', en: 'Zoom out' },
  'shell.statusBar.zoomLevel': { pt: 'Zoom', en: 'Zoom' },
  'shell.statusBar.zoomFitWidth': { pt: 'Ajustar à largura', en: 'Fit to width' },
  'shell.statusBar.notSavedYet': { pt: 'Arquivo ainda não salvo', en: 'File not saved yet' },
  'shell.statusBar.pages': {
    pt: { one: '{count} página', other: '{count} páginas' },
    en: { one: '{count} page', other: '{count} pages' },
  },
  'shell.statusBar.pageOf': { pt: 'Página {page} de {pages}', en: 'Page {page} of {pages}' },
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

  'shell.home.subtitle': {
    pt: 'Documentos e planilhas, sem depender de nuvem.',
    en: 'Documents and spreadsheets, without depending on the cloud.',
  },
  'shell.home.newDocument': { pt: 'Novo documento', en: 'New document' },
  'shell.home.newSpreadsheet': { pt: 'Nova planilha', en: 'New spreadsheet' },
  'shell.home.openFile': { pt: 'Abrir arquivo', en: 'Open file' },
  'shell.home.fromTemplate': { pt: 'A partir de modelo', en: 'From template' },

  'shell.template.title': { pt: 'Novo a partir de modelo', en: 'New from template' },
  'shell.template.builtin': { pt: 'Modelos do Librevia', en: 'Librevia templates' },
  'shell.template.user': { pt: 'Meus modelos', en: 'My templates' },
  'shell.template.userEmpty': {
    pt: 'Nenhum modelo na pasta ainda. Salve um documento como “Modelo do Word (.dotx)” nela.',
    en: 'No templates in the folder yet. Save a document as “Word template (.dotx)” there.',
  },
  'shell.template.userTemplate': { pt: 'Modelo do Word', en: 'Word template' },
  'shell.template.create': { pt: 'Criar', en: 'Create' },
  'shell.template.browse': { pt: 'Procurar…', en: 'Browse…' },
  'shell.template.openFolder': { pt: 'Abrir pasta de modelos', en: 'Open templates folder' },
  'shell.template.cancel': { pt: 'Cancelar', en: 'Cancel' },
  'shell.template.loading': { pt: 'Carregando modelos…', en: 'Loading templates…' },
  'shell.template.blank.name': { pt: 'Documento em branco', en: 'Blank document' },
  'shell.template.blank.description': {
    pt: 'Página A4 com os estilos padrão: Calibri 11, títulos e listas.',
    en: 'A4 page with the default styles: Calibri 11, headings and lists.',
  },
  'shell.template.letter.name': { pt: 'Carta', en: 'Letter' },
  'shell.template.letter.description': {
    pt: 'Remetente, data, destinatário, saudação e assinatura, com rodapé de contato.',
    en: 'Sender, date, recipient, greeting and signature, with a contact footer.',
  },
  'shell.template.report.name': { pt: 'Relatório com capa e sumário', en: 'Report with cover and contents' },
  'shell.template.report.description': {
    pt: 'Capa, sumário, seções numeradas por título e número de página no rodapé.',
    en: 'Cover page, table of contents, heading sections and page numbers in the footer.',
  },
  'shell.template.minutes.name': { pt: 'Ata de reunião', en: 'Meeting minutes' },
  'shell.template.minutes.description': {
    pt: 'Data, participantes, pauta, deliberações e tabela de encaminhamentos.',
    en: 'Date, attendees, agenda, decisions and a table of action items.',
  },

  'shell.home.recentFiles': { pt: 'Arquivos recentes', en: 'Recent files' },
  'shell.home.clear': { pt: 'Limpar', en: 'Clear' },
  'shell.home.emptyRecents': { pt: 'Nenhum arquivo aberto ainda.', en: 'No files opened yet.' },

  'shell.error.unexpected': {
    pt: 'Ocorreu um erro inesperado. A operação não foi concluída.',
    en: 'An unexpected error occurred. The operation could not be completed.',
  },
  'shell.error.communicationFailed': {
    pt: 'A comunicação com o aplicativo falhou. Tente novamente.',
    en: 'Communication with the application failed. Please try again.',
  },
  'shell.failure.main.title': { pt: 'Erro interno', en: 'Internal error' },
  'shell.failure.main.detail': {
    pt: 'Algo falhou dentro do {app}. Salve seu trabalho e abra o aplicativo de novo. Os detalhes ficaram em:\n{log}',
    en: 'Something failed inside {app}. Save your work and open the application again. The details were written to:\n{log}',
  },
  'shell.failure.window.title': { pt: 'A janela parou de funcionar.', en: 'The window stopped working.' },
  'shell.failure.window.detail': {
    pt: 'Ao reabrir, se havia alterações não salvas, o {app} oferece o rascunho guardado. Os detalhes ficaram em:\n{log}',
    en: 'When it reopens, if there were unsaved changes, {app} offers the stored draft. The details were written to:\n{log}',
  },
  'shell.failure.window.reload': { pt: 'Reabrir', en: 'Reopen' },
  'shell.failure.window.close': { pt: 'Fechar', en: 'Close' },
  'shell.failure.screen.title': { pt: 'A tela parou de funcionar.', en: 'The screen stopped working.' },
  'shell.failure.screen.detail': {
    pt: 'Ao reabrir, se havia alterações não salvas, o {app} oferece o rascunho guardado. Os detalhes ficam em Ajuda → Abrir pasta de registros.',
    en: 'When it reopens, if there were unsaved changes, {app} offers the stored draft. The details are in Help → Open log folder.',
  },
  'shell.error.rootNotFound': {
    pt: 'Elemento #root não encontrado.',
    en: '#root element not found.',
  },
  'shell.draft.autosaveBroken': {
    pt: 'A gravação automática de segurança parou de funcionar. Seu arquivo não foi alterado, mas salve o trabalho manualmente: uma queda agora custaria o que não foi salvo.',
    en: 'Automatic backup saving stopped working. Your file was not modified, but save your work manually: a crash now would cost what was not saved.',
  },
  'shell.pdf.viewer': { pt: 'Visualizador de PDF', en: 'PDF viewer' },
  'shell.pdf.viewOnly': {
    pt: 'O PDF abre só para leitura; em Arquivo → Propriedades dá para alterar título, autor, assunto e palavras-chave. Para imprimir ou salvar uma cópia, use Arquivo → Imprimir.',
    en: 'The PDF opens for reading only; File → Properties changes its title, author, subject and keywords. To print or save a copy, use File → Print.',
  },
  'shell.pdf.propertiesTitle': { pt: 'Propriedades do PDF', en: 'PDF properties' },
  'shell.pdf.application': { pt: 'Aplicativo', en: 'Application' },
  'shell.pdf.producer': { pt: 'Produtor do PDF', en: 'PDF producer' },
  'shell.pdf.version': { pt: 'Versão do PDF', en: 'PDF version' },
  'shell.pdf.signedWarning': {
    pt: 'Este PDF é assinado digitalmente. Salvar acrescenta as propriedades no fim do arquivo: a assinatura continua valendo para a versão assinada, mas leitores como o Acrobat vão avisar que o arquivo mudou depois dela.',
    en: 'This PDF is digitally signed. Saving appends the properties to the end of the file: the signature stays valid for the signed version, but readers such as Acrobat will warn that the file changed after it.',
  },
  'shell.pdf.encryptedProperties': {
    pt: 'Este PDF é criptografado. O Librevia não altera as propriedades de PDFs criptografados.',
    en: 'This PDF is encrypted. Librevia does not change the properties of encrypted PDFs.',
  },
  'shell.pdf.password': {
    pt: 'Este PDF é protegido por senha. Para lê-lo, abra o arquivo de novo e digite a senha.',
    en: 'This PDF is password protected. To read it, open the file again and type the password.',
  },
  'shell.pdf.passwordTitle': { pt: 'PDF protegido por senha', en: 'Password-protected PDF' },
  'shell.pdf.passwordLabel': { pt: 'Senha', en: 'Password' },
  'shell.pdf.passwordHint': {
    pt: 'A senha serve só para mostrar este PDF: o Librevia não a guarda.',
    en: 'The password is used only to show this PDF: Librevia does not keep it.',
  },
  'shell.pdf.passwordWrong': { pt: 'Senha incorreta. Tente de novo.', en: 'Wrong password. Try again.' },
  'shell.pdf.passwordOpen': { pt: 'Abrir', en: 'Open' },
  'shell.pdf.invalid': {
    pt: 'Não foi possível ler este PDF: o arquivo está danificado ou não é um PDF.',
    en: 'This PDF could not be read: the file is damaged or is not a PDF.',
  },
  'shell.pdf.failed': { pt: 'Não foi possível mostrar este PDF.', en: 'This PDF could not be shown.' },
  'shell.print.nothingToPrint': {
    pt: 'Não há nada aberto para imprimir. Abra ou crie um documento ou uma planilha primeiro.',
    en: 'Nothing is open to print. Open or create a document or spreadsheet first.',
  },
  'shell.export.documentOnly': {
    pt: 'Só documentos podem ser exportados como HTML, Markdown ou ODT. Abra ou crie um documento primeiro.',
    en: 'Only documents can be exported as HTML, Markdown or ODT. Open or create a document first.',
  },
  'shell.print.defaultDocumentName': {
    pt: 'Documento',
    en: 'Document',
  },
} satisfies Catalog
