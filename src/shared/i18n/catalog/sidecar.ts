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
  'sidecar.notice.comments': { pt: 'comentários', en: 'comments' },
  'sidecar.notice.structureRevisions': { pt: 'revisões de estrutura', en: 'structure revisions' },
  'sidecar.notice.formatRevisions': { pt: 'revisões de formatação', en: 'formatting revisions' },
  'sidecar.notice.footnotes': { pt: 'notas de rodapé', en: 'footnotes' },
  'sidecar.notice.endnotes': { pt: 'notas de fim', en: 'endnotes' },
  'sidecar.notice.fields': {
    pt: 'campos calculados (como sumário e número de página)',
    en: 'calculated fields (such as table of contents and page number)',
  },
  'sidecar.notice.headerFields': {
    pt: 'campos calculados no cabeçalho',
    en: 'calculated fields in the header',
  },
  'sidecar.notice.shapes': { pt: 'moldura e preenchimento de formas', en: 'shape outline and fill' },
  'sidecar.notice.contentControls': { pt: 'controles de conteúdo', en: 'content controls' },
  'sidecar.notice.equations': { pt: 'equações', en: 'equations' },
  'sidecar.notice.drawings': { pt: 'desenhos', en: 'drawings' },
  'sidecar.notice.smartTags': { pt: 'marcações inteligentes', en: 'smart tags' },
  'sidecar.notice.element': { pt: '{value}', en: '{value}' },
  'sidecar.notice.foreignBands': {
    pt: 'cabeçalho e rodapé do arquivo .docx de origem',
    en: 'header and footer from the original .docx file',
  },
  'sidecar.notice.headerBoxFieldText': {
    pt: 'texto de uma caixa de cabeçalho com campo calculado',
    en: 'text of a header box with a calculated field',
  },
  'sidecar.notice.headerBoxBookmark': {
    pt: 'marcador numa caixa de cabeçalho que você editou',
    en: 'bookmark in a header box you edited',
  },
  'sidecar.notice.headerBoxRevisions': {
    pt: 'revisões numa caixa de cabeçalho que você editou',
    en: 'revisions in a header box you edited',
  },
  'sidecar.notice.unlinkedBand': {
    pt: 'cabeçalho ou rodapé desvinculado da seção anterior',
    en: 'header or footer unlinked from the previous section',
  },
  'sidecar.notice.plainHeaderKept': {
    pt: 'cabeçalho de texto simples: o documento já tem um cabeçalho próprio',
    en: 'plain text header: the document already has a header of its own',
  },
  'sidecar.notice.plainFooterKept': {
    pt: 'rodapé de texto simples: o documento já tem um rodapé próprio',
    en: 'plain text footer: the document already has a footer of its own',
  },
  'sidecar.notice.editedCommentFormatting': {
    pt: 'formatação de um comentário que você editou',
    en: 'formatting of a comment you edited',
  },
  'sidecar.notice.templateMacros': { pt: 'macros (VBA) do modelo', en: 'template macros (VBA)' },
  'sidecar.notice.looseContent': {
    pt: 'conteúdo solto entre blocos que você apagou',
    en: 'loose content between blocks you deleted',
  },
  'sidecar.notice.editedParagraphComment': {
    pt: 'comentário ancorado num parágrafo que você editou',
    en: 'comment anchored in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphRevisionMarks': {
    pt: 'marcas de revisão num parágrafo que você editou',
    en: 'revision marks in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphRevisedImage': {
    pt: 'imagem dentro de uma revisão num parágrafo que você editou (ficou como aceita)',
    en: 'image inside a revision in a paragraph you edited (kept as accepted)',
  },
  'sidecar.notice.editedParagraphFormatRevision': {
    pt: 'revisão de formatação num parágrafo que você editou',
    en: 'formatting revision in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphMove': {
    pt: 'movimentação de texto num parágrafo que você editou (virou exclusão e inserção)',
    en: 'text move in a paragraph you edited (became a deletion and an insertion)',
  },
  'sidecar.notice.editedParagraphFootnote': {
    pt: 'nota de rodapé num parágrafo que você editou',
    en: 'footnote in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphEndnote': {
    pt: 'nota de fim num parágrafo que você editou',
    en: 'endnote in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphEquation': {
    pt: 'equação num parágrafo que você editou',
    en: 'equation in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphField': {
    pt: 'campo calculado num parágrafo que você editou',
    en: 'calculated field in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphShape': {
    pt: 'forma ou caixa de texto num parágrafo que você editou',
    en: 'shape or text box in a paragraph you edited',
  },
  'sidecar.notice.editedParagraphBoxText': {
    pt: 'texto de caixa num parágrafo que você editou',
    en: 'text box text in a paragraph you edited',
  },
  'sidecar.notice.imageUnsupported': {
    pt: 'imagem em formato não suportado',
    en: 'image in an unsupported format',
  },
  'sidecar.notice.imageWithoutContent': {
    pt: 'imagem sem conteúdo embutido',
    en: 'image without embedded content',
  },
  'sidecar.notice.imageInvalidAddress': {
    pt: 'imagem com endereço inválido',
    en: 'image with an invalid address',
  },
  'sidecar.notice.imageUnreadable': {
    pt: 'imagem com conteúdo ilegível',
    en: 'image with unreadable content',
  },
  'sidecar.notice.imageType': { pt: 'imagem em {value}', en: 'image in {value}' },
  'sidecar.notice.listNumberFormatting': {
    pt: 'formatação própria do número de uma lista (fonte, cor ou estilo do nível)',
    en: 'own formatting of a list number (font, color or level style)',
  },
  'sidecar.notice.listNumberFormat': {
    pt: 'formato de numeração de lista "{value}" (gravado como decimal)',
    en: 'list numbering format "{value}" (saved as decimal)',
  },
  'sidecar.notice.differentFirstPageRefused': {
    pt: '"Primeira página diferente" (o arquivo não aceitou o interruptor)',
    en: '"Different first page" (the file did not accept the switch)',
  },
  'sidecar.notice.differentOddEvenRefused': {
    pt: '"Pares e ímpares diferentes" (o arquivo não aceitou o interruptor)',
    en: '"Different odd and even pages" (the file did not accept the switch)',
  },
  'sidecar.notice.pageNumberingRefused': {
    pt: 'formato e início da numeração de página (o arquivo não aceitou a mudança)',
    en: 'page number format and start (the file did not accept the change)',
  },
  'sidecar.notice.trackChangesRefused': {
    pt: '"Controlar alterações" (o arquivo não aceitou o interruptor)',
    en: '"Track changes" (the file did not accept the switch)',
  },
  'sidecar.notice.pageNumberFormat': {
    pt: 'formato de número de página "{value}" (mostrado em algarismos; o arquivo o mantém)',
    en: 'page number format "{value}" (shown as digits; the file keeps it)',
  },
  'sidecar.notice.paperSize': {
    pt: 'o tamanho do papel deste documento não é A4 nem Carta (ele é preservado no arquivo)',
    en: 'the paper size of this document is neither A4 nor Letter (the file keeps it)',
  },
  'sidecar.notice.unequalColumns': {
    pt: 'colunas de larguras diferentes (mostradas iguais; o arquivo as mantém)',
    en: 'columns of different widths (shown equal; the file keeps them)',
  },
  'sidecar.notice.sectionWithoutSetup': {
    pt: 'quebra de seção sem configuração no documento (a seção foi unida à seguinte)',
    en: 'section break without setup in the document (the section was joined to the next one)',
  },
  'sidecar.notice.undefinedStyle': {
    pt: 'estilo "{value}", que o documento não define',
    en: 'style "{value}", which the document does not define',
  },
  'sidecar.notice.oldDraftStyleChange': {
    pt: 'mudança nos estilos existentes feita num rascunho de versão antiga',
    en: 'change to the existing styles made in a draft from an older version',
  },
  'sidecar.notice.staleStyleEffects': {
    pt: 'estilos com efeitos do Word 2010 (stylesWithEffects.xml)',
    en: 'Word 2010 styles with effects (stylesWithEffects.xml)',
  },
  'sidecar.notice.builtinStyleRename': {
    pt: 'novo nome do estilo embutido "{value}"',
    en: 'new name of the built-in style "{value}"',
  },
  'sidecar.notice.lineSpacing': { pt: 'entrelinha "{value}"', en: 'line spacing "{value}"' },
  'sidecar.notice.runLineSpacing': {
    pt: 'entrelinha de um trecho de texto ("{value}")',
    en: 'line spacing of a stretch of text ("{value}")',
  },
  'sidecar.notice.backgroundColor': { pt: 'cor de fundo "{value}"', en: 'background color "{value}"' },
  'sidecar.notice.fontSize': { pt: 'tamanho de fonte "{value}"', en: 'font size "{value}"' },
  'sidecar.notice.textColor': { pt: 'cor de texto "{value}"', en: 'text color "{value}"' },
  'sidecar.notice.textBackground': {
    pt: 'cor de fundo de texto "{value}"',
    en: 'text background color "{value}"',
  },
  'sidecar.notice.formatting': { pt: 'formatação "{value}"', en: 'formatting "{value}"' },
  'sidecar.notice.blockType': { pt: 'bloco do tipo "{value}"', en: 'block of type "{value}"' },
  'sidecar.notice.contentType': { pt: 'conteúdo do tipo "{value}"', en: 'content of type "{value}"' },
  'sidecar.notice.unsavedEquation': {
    pt: 'equação que não pôde ser gravada',
    en: 'equation that could not be saved',
  },
  'sidecar.notice.editedEquationFormatting': {
    pt: 'formatação dentro da equação editada',
    en: 'formatting inside the edited equation',
  },
  'sidecar.notice.invalidLink': { pt: 'endereço de link inválido', en: 'invalid link address' },
  'sidecar.notice.partialGridWidth': {
    pt: 'largura de coluna de uma tabela com linhas que não ocupam todas as colunas',
    en: 'column width of a table whose rows do not span every column',
  },
  'sidecar.notice.cellShadingPattern': {
    pt: 'trama de sombreamento de uma célula que você formatou',
    en: 'shading pattern of a cell you formatted',
  },
  'sidecar.notice.cellBorderStyle': {
    pt: 'estilo de borda de uma célula que você formatou',
    en: 'border style of a cell you formatted',
  },
  'sidecar.notice.shiftedTablePart': {
    pt: 'largura, mesclagem ou sombreamento de parte de uma tabela que você editou',
    en: 'width, merge or shading of part of a table you edited',
  },
  'sidecar.notice.editorVerticalMerge': {
    pt: 'mesclagem vertical de células feita no editor',
    en: 'vertical cell merge made in the editor',
  },
  'sidecar.notice.editedCellSpecialContent': {
    pt: 'conteúdo especial de uma célula que você editou',
    en: 'special content of a cell you edited',
  },
  'sidecar.notice.sheetTooLarge': {
    pt: 'a aba "{value}" tem mais células do que o aplicativo abre; o excedente não foi carregado',
    en: 'the sheet "{value}" has more cells than the app opens; the rest was not loaded',
  },
  'sidecar.notice.hiddenSheet': {
    pt: 'a aba "{value}" está oculta no arquivo e aparece aqui como as outras',
    en: 'the sheet "{value}" is hidden in the file and shows here like the others',
  },
  'sidecar.notice.mergedCells': {
    pt: 'células mescladas (aparecem separadas na tela)',
    en: 'merged cells (shown separate on screen)',
  },
  'sidecar.notice.conditionalFormatting': { pt: 'formatação condicional', en: 'conditional formatting' },
  'sidecar.notice.dataValidation': { pt: 'listas e validação de dados', en: 'lists and data validation' },
  'sidecar.notice.images': { pt: 'imagens', en: 'images' },
  'sidecar.notice.formattedTables': { pt: 'tabelas formatadas', en: 'formatted tables' },
  'sidecar.notice.filters': { pt: 'filtros', en: 'filters' },
  'sidecar.notice.cellComments': { pt: 'comentários nas células', en: 'cell comments' },
  'sidecar.notice.namedRanges': {
    pt: 'intervalos nomeados (fórmulas que os usam não são calculadas)',
    en: 'named ranges (formulas that use them are not calculated)',
  },
} satisfies Catalog
