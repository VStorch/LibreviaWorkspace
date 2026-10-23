import type { Catalog } from '../message.js'

/**
 * What the format service reports. It sends the key, never the sentence: the same file opens in
 * either language, and `src/main/sidecar/messages.test.ts` checks every key it uses is here.
 */
export const SIDECAR = {
  'sidecar.error.pageMissing': {
    pt: 'A configuração de página chegou vazia.',
    en: 'The page setup arrived empty.',
  },
  'sidecar.error.documentMissing': {
    pt: 'O documento a gravar chegou vazio.',
    en: 'The document to save arrived empty.',
  },
  'sidecar.error.docxWithoutDocument': {
    pt: 'O arquivo não contém um documento do Word.',
    en: 'The file does not contain a Word document.',
  },
  'sidecar.error.docxDamaged': {
    pt: 'O documento do Word está vazio ou danificado.',
    en: 'The Word document is empty or damaged.',
  },
  'sidecar.error.docxUnreadable': {
    pt: 'Não foi possível abrir este arquivo. Ele pode estar danificado ou não ser um documento do Word.',
    en: 'This file could not be opened. It may be damaged or not be a Word document.',
  },
  'sidecar.error.originalDamaged': {
    pt: 'O arquivo original está vazio ou danificado.',
    en: 'The original file is empty or damaged.',
  },
  'sidecar.error.originalWithoutDocument': {
    pt: 'O arquivo original não contém um documento do Word.',
    en: 'The original file does not contain a Word document.',
  },
  'sidecar.error.originalUnwritable': {
    pt: 'Não foi possível gravar sobre o arquivo original. Ele pode ter sido alterado ou danificado.',
    en: 'The original file could not be saved over. It may have been changed or damaged.',
  },
  'sidecar.error.noSheets': {
    pt: 'Esta planilha não tem nenhuma aba.',
    en: 'This spreadsheet has no sheets.',
  },
  'sidecar.error.nothingToSave': {
    pt: 'Não há nada para gravar: a planilha ficou sem abas.',
    en: 'There is nothing to save: the spreadsheet has no sheets left.',
  },
  'sidecar.error.workbookMissing': {
    pt: 'A planilha a gravar chegou vazia.',
    en: 'The spreadsheet to save arrived empty.',
  },
  'sidecar.error.xlsxUnreadable': {
    pt: 'Esta planilha não pôde ser lida: o arquivo parece corrompido ou não é uma planilha do Excel.',
    en: 'This spreadsheet could not be read: the file looks corrupted or is not an Excel spreadsheet.',
  },
  'sidecar.error.xlsxOriginalUnreadable': {
    pt: 'O arquivo original não pôde ser lido para gravar por cima. Use "Salvar como" para gravar num arquivo novo.',
    en: 'The original file could not be read to save over it. Use "Save as" to save to a new file.',
  },
  'sidecar.error.notPdf': { pt: 'Este arquivo não é um PDF.', en: 'This file is not a PDF.' },
  'sidecar.error.pdfDamaged': {
    pt: 'Este PDF está danificado: não foi possível ler a estrutura dele.',
    en: 'This PDF is damaged: its structure could not be read.',
  },
  'sidecar.error.pdfCompression': {
    pt: 'Este PDF usa uma compactação que o {app} não sabe ler.',
    en: 'This PDF uses a compression {app} cannot read.',
  },
  'sidecar.error.pdfEncrypted': {
    pt: 'Este PDF é criptografado. O {app} não altera as propriedades de PDFs criptografados.',
    en: 'This PDF is encrypted. {app} does not change the properties of encrypted PDFs.',
  },
  'sidecar.error.xmpDamaged': {
    pt: 'Os metadados XMP deste PDF estão danificados.',
    en: 'The XMP metadata of this PDF is damaged.',
  },
  'sidecar.error.xmpUnsupported': {
    pt: 'Os metadados XMP deste PDF estão num formato que o {app} não sabe alterar.',
    en: 'The XMP metadata of this PDF is in a format {app} cannot change.',
  },
  'sidecar.error.propertiesMissing': {
    pt: 'As propriedades a gravar chegaram vazias.',
    en: 'The properties to save arrived empty.',
  },
  'sidecar.error.notOfficeDocument': {
    pt: 'O arquivo não é um documento do Word nem uma planilha do Excel.',
    en: 'The file is neither a Word document nor an Excel spreadsheet.',
  },
  'sidecar.error.invalidPackage': {
    pt: 'O arquivo não é um pacote do Office válido.',
    en: 'The file is not a valid Office package.',
  },
  'sidecar.error.inspectionMissing': {
    pt: 'O pedido de inspeção chegou vazio.',
    en: 'The inspection request arrived empty.',
  },
  'sidecar.error.unknownMethod': {
    pt: 'O aplicativo pediu uma operação que este serviço não conhece.',
    en: 'The application asked for an operation this service does not know.',
  },
  'sidecar.error.unexpected': {
    pt: 'Ocorreu um erro inesperado ao processar o documento.',
    en: 'An unexpected error occurred while processing the document.',
  },
} satisfies Catalog
