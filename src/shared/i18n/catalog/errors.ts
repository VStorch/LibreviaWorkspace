import type { Catalog } from '../message.js'

/**
 * Frases de erro: erros de arquivo, validação de caminhos, limites,
 * serviço de formatos (sidecar), IPC e falhas de leitura/escrita.
 */
export const ERRORS = {
  // Erro genérico
  'errors.unexpected': {
    pt: 'Ocorreu um erro inesperado. A operação não foi concluída.',
    en: 'An unexpected error occurred. The operation could not be completed.',
  },

  // Sistema de arquivos (fromFileSystemError)
  'errors.fs.fileNotFound': {
    pt: 'O arquivo não foi encontrado. Ele pode ter sido movido ou excluído.',
    en: 'The file was not found. It may have been moved or deleted.',
  },
  'errors.fs.permissionDenied': {
    pt: 'Você não tem permissão para acessar este arquivo. Verifique com quem administra a pasta.',
    en: 'You do not have permission to access this file. Check with whoever manages the folder.',
  },
  'errors.fs.notAFile': {
    pt: 'O caminho indicado é uma pasta, não um arquivo.',
    en: 'The specified path is a folder, not a file.',
  },
  'errors.fs.readOnlyLocation': {
    pt: 'Este local é somente leitura. Salve o arquivo em outra pasta.',
    en: 'This location is read-only. Save the file to another folder.',
  },
  'errors.fs.diskFull': {
    pt: 'Não há espaço em disco para salvar o arquivo.',
    en: 'There is not enough disk space to save the file.',
  },
  'errors.fs.quotaExceeded': {
    pt: 'Sua cota de espaço nesta pasta de rede acabou. Libere espaço ou salve em outro lugar.',
    en: 'Your space quota on this network folder has run out. Free up space or save elsewhere.',
  },
  'errors.fs.nameTooLong': {
    pt: 'O nome do arquivo, junto com o caminho da pasta, ficou longo demais. Use um nome mais curto.',
    en: 'The file name, along with the folder path, is too long. Use a shorter name.',
  },
  'errors.fs.fileInUse': {
    pt: 'O arquivo está em uso por outro programa. Feche-o e tente novamente.',
    en: 'The file is in use by another program. Close it and try again.',
  },
  'errors.fs.networkTimeout': {
    pt: 'A pasta de rede não respondeu. Verifique a conexão e tente novamente.',
    en: 'The network folder did not respond. Check the connection and try again.',
  },
  'errors.fs.readFailed': {
    pt: 'Não foi possível ler o arquivo.',
    en: 'Could not read the file.',
  },
  'errors.fs.writeFailed': {
    pt: 'Não foi possível salvar o arquivo. O conteúdo original foi preservado.',
    en: 'Could not save the file. The original content was preserved.',
  },

  // Documento .sdoc
  'errors.document.corrupt': {
    pt: 'Este arquivo não pôde ser lido: o conteúdo está corrompido ou não é um documento válido.',
    en: 'This file could not be read: the content is corrupt or not a valid document.',
  },
  'errors.document.invalid': {
    pt: 'Este arquivo não é um documento válido deste aplicativo.',
    en: 'This file is not a valid document for this application.',
  },
  'errors.document.newerVersion': {
    pt: 'Este documento foi criado por uma versão mais recente do aplicativo. Atualize para abri-lo.',
    en: 'This document was created by a newer version of the application. Update to open it.',
  },

  // Caminhos e limites (src/main/fs/paths.ts)
  'errors.paths.unauthorized': {
    pt: 'Esta operação foi recusada porque o arquivo não foi aberto nem escolhido por você nesta sessão.',
    en: 'This operation was refused because the file was not opened or chosen by you in this session.',
  },
  'errors.paths.invalidPath': {
    pt: 'O caminho do arquivo é inválido.',
    en: 'The file path is invalid.',
  },
  'errors.paths.unsupportedType': {
    pt: 'Este tipo de arquivo não é suportado. O aplicativo abre .sdoc, .ssheet, .docx, .xlsx e .txt.',
    en: 'This file type is not supported. The application opens .sdoc, .ssheet, .docx, .xlsx, and .txt.',
  },
  'errors.paths.notAFile': {
    pt: 'O caminho indicado não é um arquivo.',
    en: 'The specified path is not a file.',
  },
  'errors.paths.fileTooLarge': {
    pt: 'O arquivo é grande demais para ser aberto (limite atual: {limit} MB).',
    en: 'The file is too large to open (current limit: {limit} MB).',
  },

  // Imagens (src/main/fs/read-image.ts)
  'errors.image.imageTooLarge': {
    pt: 'A imagem é grande demais para ser inserida (limite: {limit} MB).',
    en: 'The image is too large to insert (limit: {limit} MB).',
  },
  'errors.image.unsupportedImage': {
    pt: 'Este arquivo não é uma imagem suportada. Use PNG, JPEG, GIF ou WebP.',
    en: 'This file is not a supported image. Use PNG, JPEG, GIF, or WebP.',
  },

  // Texto (src/main/fs/read-text.ts)
  'errors.text.notTextFile': {
    pt: 'Este arquivo não parece ser de texto e não pode ser aberto com segurança.',
    en: 'This file does not appear to be text and cannot be opened safely.',
  },

  // Recuperação (src/main/fs/recovery.ts)
  'errors.recovery.folderNotConfigured': {
    pt: 'a pasta de recuperação não foi configurada',
    en: 'recovery folder has not been configured',
  },

  // IPC (src/main/ipc/)
  'errors.ipc.windowNotAvailable': {
    pt: 'A janela do aplicativo não está disponível.',
    en: 'The application window is not available.',
  },
  'errors.ipc.notInRecents': {
    pt: 'Este arquivo não está mais na lista de recentes. Abra-o novamente pelo menu Arquivo.',
    en: 'This file is no longer in the recent files list. Open it again from the File menu.',
  },
  'errors.ipc.invalidData': {
    pt: 'A operação foi recusada porque os dados enviados são inválidos.',
    en: 'The operation was refused because the provided data is invalid.',
  },
  'errors.ipc.unrecognizedEnd': {
    pt: 'A operação terminou de um jeito que o aplicativo não reconhece.',
    en: 'The operation ended in a way the application does not recognize.',
  },

  // Impressão (src/main/print/pdf.ts)
  'errors.print.prepareFailed': {
    pt: 'Não foi possível preparar o documento: {description}',
    en: 'Could not prepare document: {description}',
  },
  'errors.print.printFailed': {
    pt: 'Não foi possível imprimir: {reason}',
    en: 'Could not print: {reason}',
  },
  'errors.print.previewTitle': {
    pt: 'Visualizar impressão — {title}',
    en: 'Print preview — {title}',
  },

  // DOCX (src/main/docx/index.ts)
  'errors.docx.cannotRead': {
    pt: 'Não foi possível ler este documento do Word. O arquivo pode estar danificado.',
    en: 'Could not read this Word document. The file may be damaged.',
  },
  'errors.docx.openContract': {
    pt: 'docx.open fora do contrato',
    en: 'docx.open out of contract',
  },
  'errors.docx.saveOnlyFromDocx': {
    pt: 'Só é possível salvar em .docx um documento que foi aberto a partir de um arquivo .docx. Salve como .sdoc ou abra um documento do Word primeiro.',
    en: 'You can only save to .docx a document that was opened from a .docx file. Save as .sdoc or open a Word document first.',
  },
  'errors.docx.cannotSave': {
    pt: 'Não foi possível gravar o documento do Word. O arquivo original não foi alterado.',
    en: 'Could not save the Word document. The original file was not modified.',
  },
  'errors.docx.saveContract': {
    pt: 'docx.save fora do contrato',
    en: 'docx.save out of contract',
  },
  'errors.docx.inconsistentState': {
    pt: 'O documento em edição está em estado inconsistente.',
    en: 'The document being edited is in an inconsistent state.',
  },
  'errors.docx.foreignBands': {
    pt: 'cabeçalho e rodapé do arquivo .docx de origem',
    en: 'header and footer from the original .docx file',
  },
  'errors.docx.originPackage': {
    pt: 'estilos, notas, comentários e demais partes do arquivo .docx de origem',
    en: 'styles, notes, comments, and other parts of the original .docx file',
  },
  'errors.docx.cannotCreate': {
    pt: 'Não foi possível criar o documento do Word. Nada foi gravado.',
    en: 'Could not create the Word document. Nothing was saved.',
  },
  'errors.docx.createContract': { pt: 'docx.create sem pacote', en: 'docx.create returned no package' },

  // XLSX (src/main/xlsx/index.ts)
  'errors.xlsx.cannotRead': {
    pt: 'Não foi possível ler esta planilha do Excel. O arquivo pode estar danificado.',
    en: 'Could not read this Excel spreadsheet. The file may be damaged.',
  },
  'errors.xlsx.openContract': {
    pt: 'xlsx.open fora do contrato',
    en: 'xlsx.open out of contract',
  },
  'errors.xlsx.cannotSave': {
    pt: 'Não foi possível gravar a planilha do Excel. O arquivo original não foi alterado.',
    en: 'Could not save the Excel spreadsheet. The original file was not modified.',
  },
  'errors.xlsx.saveContract': {
    pt: 'xlsx.save fora do contrato',
    en: 'xlsx.save out of contract',
  },
  'errors.xlsx.invalidSchema': {
    pt: 'xlsx.open devolveu um modelo fora do esquema',
    en: 'xlsx.open returned a model out of schema',
  },
  'errors.xlsx.onlySpreadsheets': {
    pt: 'Só é possível salvar em .xlsx uma planilha. Para um documento de texto, use .docx ou .sdoc.',
    en: 'You can only save a spreadsheet to .xlsx. For a text document, use .docx or .sdoc.',
  },
  'errors.xlsx.unknownFunctions': {
    pt: 'funções que este aplicativo não calcula: {names}',
    en: 'functions this application does not calculate: {names}',
  },

  // Sidecar (src/main/sidecar/)
  'errors.sidecar.died': {
    pt: 'O serviço de formatos foi encerrado inesperadamente. Seu documento continua aberto e intacto.',
    en: 'The format service terminated unexpectedly. Your document remains open and intact.',
  },
  'errors.sidecar.timedOut': {
    pt: 'O serviço de formatos demorou demais para responder e a operação foi cancelada. Seu documento continua aberto e intacto.',
    en: 'The format service took too long to respond and the operation was canceled. Your document remains open and intact.',
  },
  'errors.sidecar.healthContract': {
    pt: 'health fora do contrato',
    en: 'health out of contract',
  },
  'errors.sidecar.alreadyClosed': {
    pt: 'cliente já encerrado',
    en: 'client already closed',
  },
  'errors.sidecar.closedDuringRequest': {
    pt: 'encerrado durante o pedido',
    en: 'closed during request',
  },
  'errors.sidecar.exitCodeSignal': {
    pt: 'saiu com code={code} signal={signal}',
    en: 'exited with code={code} signal={signal}',
  },
  'errors.sidecar.serviceNotFound': {
    pt: 'O serviço que lê e grava documentos do Office não foi encontrado. A instalação parece incompleta — reinstale o aplicativo.',
    en: 'The service that reads and writes Office documents was not found. The installation seems incomplete — reinstall the application.',
  },
  'errors.sidecar.noBinary': {
    pt: 'sem binário publicado para {platform}-{arch}',
    en: 'no published binary for {platform}-{arch}',
  },
  'errors.sidecar.missingOrNotExecutable': {
    pt: 'ausente ou sem permissão de execução',
    en: 'missing or no execution permission',
  },
  'errors.sidecar.unexpectedResponse': {
    pt: 'O serviço de formatos respondeu de forma inesperada. A operação não foi concluída.',
    en: 'The format service responded unexpectedly. The operation could not be completed.',
  },
  'errors.sidecar.frameAnnounce': {
    pt: 'quadro anuncia {jsonLength} bytes de JSON e {binaryLength} de binário',
    en: 'frame announces {jsonLength} bytes of JSON and {binaryLength} of binary',
  },
  'errors.sidecar.invalidJson': {
    pt: 'JSON inválido no quadro',
    en: 'invalid JSON in frame',
  },
  'errors.sidecar.contractViolation': {
    pt: 'resposta fora do contrato',
    en: 'response out of contract',
  },
} satisfies Catalog
