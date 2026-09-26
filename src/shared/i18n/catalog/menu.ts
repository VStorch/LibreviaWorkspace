import type { Catalog } from '../message.js'

/**
 * O menu nativo.
 *
 * Montado no processo main, e por isso a primeira área a entrar no catálogo:
 * ela é a prova de que a tradução atravessa a fronteira do IPC. Uma solução que
 * só funcionasse no renderer deixaria a barra de menus em português para sempre.
 *
 * `{app}` é o nome do aplicativo, que não se traduz.
 */
export const MENU = {
  'menu.file': { pt: 'Arquivo', en: 'File' },
  'menu.edit': { pt: 'Editar', en: 'Edit' },
  'menu.format': { pt: 'Formatar', en: 'Format' },
  'menu.table': { pt: 'Tabela', en: 'Table' },
  'menu.insert': { pt: 'Inserir', en: 'Insert' },
  'menu.view': { pt: 'Exibir', en: 'View' },
  'menu.tools': { pt: 'Ferramentas', en: 'Tools' },
  'menu.help': { pt: 'Ajuda', en: 'Help' },

  'menu.file.newDocument': { pt: 'Novo documento', en: 'New document' },
  'menu.file.newSpreadsheet': { pt: 'Nova planilha', en: 'New spreadsheet' },
  'menu.file.open': { pt: 'Abrir…', en: 'Open…' },
  'menu.file.openRecent': { pt: 'Abrir recente', en: 'Open recent' },
  'menu.file.noRecent': { pt: 'Nenhum arquivo recente', en: 'No recent files' },
  'menu.file.clearRecent': { pt: 'Limpar recentes', en: 'Clear recent files' },
  'menu.file.save': { pt: 'Salvar', en: 'Save' },
  'menu.file.saveAs': { pt: 'Salvar como…', en: 'Save as…' },
  'menu.file.pageSetup': { pt: 'Configuração de página…', en: 'Page setup…' },
  'menu.file.printPreview': { pt: 'Visualizar impressão', en: 'Print preview' },
  'menu.file.exportPdf': { pt: 'Exportar para PDF…', en: 'Export to PDF…' },
  'menu.file.print': { pt: 'Imprimir…', en: 'Print…' },
  'menu.file.close': { pt: 'Fechar arquivo', en: 'Close file' },
  'menu.file.quit': { pt: 'Sair', en: 'Quit' },

  'menu.edit.undo': { pt: 'Desfazer', en: 'Undo' },
  'menu.edit.redo': { pt: 'Refazer', en: 'Redo' },
  'menu.edit.cut': { pt: 'Recortar', en: 'Cut' },
  'menu.edit.copy': { pt: 'Copiar', en: 'Copy' },
  'menu.edit.paste': { pt: 'Colar', en: 'Paste' },
  'menu.edit.pasteWithoutFormat': { pt: 'Colar sem formatação', en: 'Paste without formatting' },
  'menu.edit.selectAll': { pt: 'Selecionar tudo', en: 'Select all' },
  'menu.edit.findReplace': { pt: 'Localizar e substituir…', en: 'Find and replace…' },

  'menu.format.paragraph': { pt: 'Parágrafo…', en: 'Paragraph…' },
  'menu.format.image': { pt: 'Imagem…', en: 'Image…' },

  'menu.insert.pageBreak': { pt: 'Quebra de página', en: 'Page break' },
  'menu.insert.specialCharacter': { pt: 'Caractere especial…', en: 'Special character…' },
  'menu.insert.bookmark': { pt: 'Marcador…', en: 'Bookmark…' },

  'menu.view.formattingMarks': { pt: 'Marcas de formatação', en: 'Formatting marks' },
  'menu.view.navigationPane': { pt: 'Painel de navegação', en: 'Navigation pane' },
  'menu.view.resetZoom': { pt: 'Tamanho normal', en: 'Actual size' },
  'menu.view.zoomIn': { pt: 'Ampliar', en: 'Zoom in' },
  'menu.view.zoomOut': { pt: 'Reduzir', en: 'Zoom out' },
  'menu.view.zoomFitWidth': { pt: 'Ajustar à largura', en: 'Fit to width' },
  'menu.view.fullScreen': { pt: 'Tela cheia', en: 'Full screen' },
  'menu.view.reload': { pt: 'Recarregar', en: 'Reload' },
  'menu.view.devTools': { pt: 'Ferramentas do desenvolvedor', en: 'Developer tools' },

  'menu.tools.spellcheck': { pt: 'Verificação ortográfica', en: 'Spell check' },
  'menu.tools.typography': { pt: 'Autocorreção tipográfica', en: 'Smart typography' },
  'menu.tools.wordCount': { pt: 'Contar palavras…', en: 'Word count…' },

  'menu.help.about': { pt: 'Sobre o {app}', en: 'About {app}' },

  // Só aparecem no macOS, e é o sistema que dita a redação destes.
  'menu.app.hide': { pt: 'Ocultar {app}', en: 'Hide {app}' },
  'menu.app.hideOthers': { pt: 'Ocultar outros', en: 'Hide others' },
  'menu.app.unhide': { pt: 'Mostrar todos', en: 'Show all' },
  'menu.app.quit': { pt: 'Encerrar {app}', en: 'Quit {app}' },
} satisfies Catalog
