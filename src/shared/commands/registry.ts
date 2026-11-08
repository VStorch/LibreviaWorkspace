import type { MessageKey } from '../i18n/index.js'
import type { ShortcutId } from '../shortcuts.js'
import { MenuCommand } from '../types.js'

export type CommandId = MenuCommand

export interface CommandMeta {
  readonly id: CommandId
  readonly labelKey: MessageKey
  readonly shortcut?: ShortcutId
  /** The page handles the key; the menu only shows it, or on macOS it would take it from the page. */
  readonly registerAccelerator?: false
}

export const COMMANDS = {
  [MenuCommand.NewDocument]: {
    id: MenuCommand.NewDocument,
    labelKey: 'menu.file.newDocument',
    shortcut: 'newDocument',
  },
  [MenuCommand.NewSpreadsheet]: {
    id: MenuCommand.NewSpreadsheet,
    labelKey: 'menu.file.newSpreadsheet',
    shortcut: 'newSpreadsheet',
  },
  [MenuCommand.NewFromTemplate]: { id: MenuCommand.NewFromTemplate, labelKey: 'menu.file.newFromTemplate' },
  [MenuCommand.Open]: { id: MenuCommand.Open, labelKey: 'menu.file.open', shortcut: 'open' },
  [MenuCommand.OpenRecent]: { id: MenuCommand.OpenRecent, labelKey: 'menu.file.openRecent' },
  [MenuCommand.ClearRecent]: { id: MenuCommand.ClearRecent, labelKey: 'menu.file.clearRecent' },
  [MenuCommand.Save]: { id: MenuCommand.Save, labelKey: 'menu.file.save', shortcut: 'save' },
  [MenuCommand.SaveAs]: { id: MenuCommand.SaveAs, labelKey: 'menu.file.saveAs', shortcut: 'saveAs' },
  [MenuCommand.CloseFile]: { id: MenuCommand.CloseFile, labelKey: 'menu.file.close', shortcut: 'closeFile' },
  [MenuCommand.Undo]: { id: MenuCommand.Undo, labelKey: 'menu.edit.undo', registerAccelerator: false },
  [MenuCommand.Redo]: { id: MenuCommand.Redo, labelKey: 'menu.edit.redo', registerAccelerator: false },
  [MenuCommand.FindReplace]: {
    id: MenuCommand.FindReplace,
    labelKey: 'menu.edit.findReplace',
    shortcut: 'findReplace',
  },
  [MenuCommand.ExportPdf]: { id: MenuCommand.ExportPdf, labelKey: 'menu.file.exportPdf' },
  [MenuCommand.ExportHtml]: { id: MenuCommand.ExportHtml, labelKey: 'menu.file.exportHtml' },
  [MenuCommand.ExportMarkdown]: { id: MenuCommand.ExportMarkdown, labelKey: 'menu.file.exportMarkdown' },
  [MenuCommand.ExportOdt]: { id: MenuCommand.ExportOdt, labelKey: 'menu.file.exportOdt' },
  [MenuCommand.ExportWord]: { id: MenuCommand.ExportWord, labelKey: 'menu.file.exportWord' },
  [MenuCommand.Print]: { id: MenuCommand.Print, labelKey: 'menu.file.print', shortcut: 'print' },
  [MenuCommand.PrintPreview]: { id: MenuCommand.PrintPreview, labelKey: 'menu.file.printPreview' },
  [MenuCommand.PageSetup]: { id: MenuCommand.PageSetup, labelKey: 'menu.file.pageSetup' },
  [MenuCommand.ParagraphSetup]: { id: MenuCommand.ParagraphSetup, labelKey: 'menu.format.paragraph' },
  [MenuCommand.InsertPageBreak]: {
    id: MenuCommand.InsertPageBreak,
    labelKey: 'menu.insert.pageBreak',
    shortcut: 'insertPageBreak',
  },
  [MenuCommand.InsertSectionNextPage]: {
    id: MenuCommand.InsertSectionNextPage,
    labelKey: 'menu.insert.sectionNextPage',
  },
  [MenuCommand.InsertSectionContinuous]: {
    id: MenuCommand.InsertSectionContinuous,
    labelKey: 'menu.insert.sectionContinuous',
  },
  [MenuCommand.InsertSectionEvenPage]: {
    id: MenuCommand.InsertSectionEvenPage,
    labelKey: 'menu.insert.sectionEvenPage',
  },
  [MenuCommand.InsertSectionOddPage]: {
    id: MenuCommand.InsertSectionOddPage,
    labelKey: 'menu.insert.sectionOddPage',
  },
  [MenuCommand.DeleteSectionBreak]: {
    id: MenuCommand.DeleteSectionBreak,
    labelKey: 'menu.insert.deleteSectionBreak',
  },
  [MenuCommand.InsertColumnBreak]: { id: MenuCommand.InsertColumnBreak, labelKey: 'menu.insert.columnBreak' },
  [MenuCommand.FormatColumns]: { id: MenuCommand.FormatColumns, labelKey: 'menu.format.columns' },
  [MenuCommand.ZoomIn]: { id: MenuCommand.ZoomIn, labelKey: 'menu.view.zoomIn', shortcut: 'zoomIn' },
  [MenuCommand.ZoomOut]: { id: MenuCommand.ZoomOut, labelKey: 'menu.view.zoomOut', shortcut: 'zoomOut' },
  [MenuCommand.ZoomReset]: {
    id: MenuCommand.ZoomReset,
    labelKey: 'menu.view.resetZoom',
    shortcut: 'zoomReset',
  },
  [MenuCommand.ZoomFitWidth]: { id: MenuCommand.ZoomFitWidth, labelKey: 'menu.view.zoomFitWidth' },
  [MenuCommand.PasteWithoutFormat]: {
    id: MenuCommand.PasteWithoutFormat,
    labelKey: 'menu.edit.pasteWithoutFormat',
    shortcut: 'pasteWithoutFormat',
    registerAccelerator: false,
  },
  [MenuCommand.WordCount]: {
    id: MenuCommand.WordCount,
    labelKey: 'menu.tools.wordCount',
    shortcut: 'wordCount',
  },
  [MenuCommand.DocumentProperties]: { id: MenuCommand.DocumentProperties, labelKey: 'menu.file.properties' },
  [MenuCommand.InspectDocument]: { id: MenuCommand.InspectDocument, labelKey: 'menu.file.inspect' },
  [MenuCommand.SpecialCharacter]: {
    id: MenuCommand.SpecialCharacter,
    labelKey: 'menu.insert.specialCharacter',
  },
  [MenuCommand.InsertEquation]: {
    id: MenuCommand.InsertEquation,
    labelKey: 'menu.insert.equation',
    shortcut: 'insertEquation',
  },
  [MenuCommand.InsertDisplayEquation]: {
    id: MenuCommand.InsertDisplayEquation,
    labelKey: 'menu.insert.displayEquation',
  },
  [MenuCommand.EditEquation]: { id: MenuCommand.EditEquation, labelKey: 'document.math.dialog.title' },
  [MenuCommand.ImageProperties]: { id: MenuCommand.ImageProperties, labelKey: 'menu.format.image' },
  [MenuCommand.InsertBookmark]: {
    id: MenuCommand.InsertBookmark,
    labelKey: 'menu.insert.bookmark',
    shortcut: 'insertBookmark',
  },
  [MenuCommand.InsertComment]: {
    id: MenuCommand.InsertComment,
    labelKey: 'menu.insert.comment',
    shortcut: 'insertComment',
  },
  [MenuCommand.InsertFootnote]: {
    id: MenuCommand.InsertFootnote,
    labelKey: 'menu.insert.footnote',
    shortcut: 'insertFootnote',
  },
  [MenuCommand.InsertEndnote]: {
    id: MenuCommand.InsertEndnote,
    labelKey: 'menu.insert.endnote',
    shortcut: 'insertEndnote',
  },
  [MenuCommand.NextComment]: { id: MenuCommand.NextComment, labelKey: 'menu.insert.nextComment' },
  [MenuCommand.PreviousComment]: { id: MenuCommand.PreviousComment, labelKey: 'menu.insert.previousComment' },
  [MenuCommand.AcceptChange]: { id: MenuCommand.AcceptChange, labelKey: 'revisions.accept' },
  [MenuCommand.RejectChange]: { id: MenuCommand.RejectChange, labelKey: 'revisions.reject' },
  [MenuCommand.AcceptAllChanges]: { id: MenuCommand.AcceptAllChanges, labelKey: 'revisions.acceptAll' },
  [MenuCommand.RejectAllChanges]: { id: MenuCommand.RejectAllChanges, labelKey: 'revisions.rejectAll' },
  [MenuCommand.NextChange]: { id: MenuCommand.NextChange, labelKey: 'revisions.next' },
  [MenuCommand.PreviousChange]: { id: MenuCommand.PreviousChange, labelKey: 'revisions.previous' },
  [MenuCommand.ToggleTrackChanges]: {
    id: MenuCommand.ToggleTrackChanges,
    labelKey: 'revisions.track',
    shortcut: 'trackChanges',
  },
  [MenuCommand.ShowAllMarkup]: { id: MenuCommand.ShowAllMarkup, labelKey: 'revisions.show.all' },
  [MenuCommand.ShowSimpleMarkup]: { id: MenuCommand.ShowSimpleMarkup, labelKey: 'revisions.show.simple' },
  [MenuCommand.ShowNoMarkup]: { id: MenuCommand.ShowNoMarkup, labelKey: 'revisions.show.none' },
  [MenuCommand.ShowOriginal]: { id: MenuCommand.ShowOriginal, labelKey: 'revisions.show.original' },
  [MenuCommand.AuthorName]: { id: MenuCommand.AuthorName, labelKey: 'menu.tools.authorName' },
  [MenuCommand.Options]: { id: MenuCommand.Options, labelKey: 'menu.tools.options' },
  [MenuCommand.InsertTableOfContents]: {
    id: MenuCommand.InsertTableOfContents,
    labelKey: 'menu.references.tableOfContents',
  },
  [MenuCommand.UpdateTableOfContents]: {
    id: MenuCommand.UpdateTableOfContents,
    labelKey: 'menu.references.updateTableOfContents',
  },
  [MenuCommand.UpdateFields]: {
    id: MenuCommand.UpdateFields,
    labelKey: 'menu.references.updateFields',
    shortcut: 'updateFields',
  },
  [MenuCommand.InsertCaption]: { id: MenuCommand.InsertCaption, labelKey: 'menu.references.caption' },
  [MenuCommand.InsertCrossReference]: {
    id: MenuCommand.InsertCrossReference,
    labelKey: 'menu.references.crossReference',
  },
  [MenuCommand.InsertPropertyField]: {
    id: MenuCommand.InsertPropertyField,
    labelKey: 'menu.insert.propertyField',
  },
  [MenuCommand.AiAssistant]: {
    id: MenuCommand.AiAssistant,
    labelKey: 'menu.tools.aiAssistant',
    shortcut: 'aiAssistant',
  },
  [MenuCommand.AiProviders]: { id: MenuCommand.AiProviders, labelKey: 'menu.tools.aiProviders' },
  [MenuCommand.TableInsert]: {
    id: MenuCommand.TableInsert,
    labelKey: 'table.insert',
    shortcut: 'insertTable',
  },
  [MenuCommand.TableRowBefore]: { id: MenuCommand.TableRowBefore, labelKey: 'table.rowBefore' },
  [MenuCommand.TableRowAfter]: { id: MenuCommand.TableRowAfter, labelKey: 'table.rowAfter' },
  [MenuCommand.TableDeleteRow]: { id: MenuCommand.TableDeleteRow, labelKey: 'table.deleteRow' },
  [MenuCommand.TableColumnBefore]: { id: MenuCommand.TableColumnBefore, labelKey: 'table.columnBefore' },
  [MenuCommand.TableColumnAfter]: { id: MenuCommand.TableColumnAfter, labelKey: 'table.columnAfter' },
  [MenuCommand.TableDeleteColumn]: { id: MenuCommand.TableDeleteColumn, labelKey: 'table.deleteColumn' },
  [MenuCommand.TableMergeCells]: { id: MenuCommand.TableMergeCells, labelKey: 'table.mergeCells' },
  [MenuCommand.TableSplitCell]: { id: MenuCommand.TableSplitCell, labelKey: 'table.splitCell' },
  [MenuCommand.TableHeaderRow]: { id: MenuCommand.TableHeaderRow, labelKey: 'table.headerRow' },
  [MenuCommand.TableDelete]: { id: MenuCommand.TableDelete, labelKey: 'table.delete' },
  [MenuCommand.TableProperties]: { id: MenuCommand.TableProperties, labelKey: 'table.properties' },
  [MenuCommand.SaveAndExit]: { id: MenuCommand.SaveAndExit, labelKey: 'dialog.discard.save' },
} as const satisfies { readonly [Id in CommandId]: CommandMeta & { readonly id: Id } }
