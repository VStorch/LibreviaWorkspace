import type { Catalog } from '../message.js'

/**
 * Aparência e modo de leitura: o que este conjunto de mudanças trouxe.
 *
 * Fica separado do menu porque as mesmas frases aparecem nos dois lugares — no
 * menu nativo e no painel de preferências da tela — e repeti-las seria repetir
 * também a decisão de como chamá-las.
 */
export const VIEW = {
  'view.showToolbar': { pt: 'Mostrar barra de ferramentas', en: 'Show toolbar' },
  'view.showStatusBar': { pt: 'Mostrar barra de status', en: 'Show status bar' },
  'view.appearance': { pt: 'Aparência', en: 'Appearance' },

  'view.theme': { pt: 'Tema', en: 'Theme' },
  'view.theme.system': { pt: 'Do sistema', en: 'System' },
  'view.theme.light': { pt: 'Claro', en: 'Light' },
  'view.theme.dark': { pt: 'Escuro', en: 'Dark' },

  'view.language': { pt: 'Idioma', en: 'Language' },

  'view.reading': { pt: 'Modo de leitura', en: 'Reading mode' },
  'view.reading.enter': { pt: 'Entrar no modo de leitura', en: 'Enter reading mode' },
  'view.reading.exit': { pt: 'Sair do modo de leitura', en: 'Exit reading mode' },
  'view.reading.hint': {
    pt: 'Modo de leitura. Esc para sair.',
    en: 'Reading mode. Press Esc to leave.',
  },
  'view.reading.readOnly': {
    pt: 'Somente leitura enquanto durar o modo de leitura.',
    en: 'Read-only for as long as reading mode lasts.',
  },

  'view.notes': { pt: 'Notas', en: 'Notes' },
  'view.notes.add': { pt: 'Adicionar nota', en: 'Add note' },
  'view.notes.edit': { pt: 'Editar nota', en: 'Edit note' },
  'view.notes.remove': { pt: 'Excluir nota', en: 'Delete note' },
  'view.notes.placeholder': { pt: 'Escreva a nota…', en: 'Write the note…' },
  'view.notes.empty': {
    pt: 'Nenhuma nota neste documento.',
    en: 'No notes in this document.',
  },
  'view.notes.count': {
    pt: { one: '{count} nota', other: '{count} notas' },
    en: { one: '{count} note', other: '{count} notes' },
  },
  // O aviso que explica por que a nota não vai parar no .docx. Aparece uma vez,
  // ao criar a primeira nota num arquivo do Office.
  'view.notes.sidecar': {
    pt: 'As notas ficam em {file}, ao lado do arquivo. O {ext} não é alterado.',
    en: 'Notes are kept in {file}, beside the file. The {ext} itself is not touched.',
  },
} satisfies Catalog
