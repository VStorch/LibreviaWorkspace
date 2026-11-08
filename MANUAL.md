# Librevia — user manual

**English** · [Português (Brasil)](MANUAL.pt-BR.md)

Document and spreadsheet editor that works **without internet**. It opens and saves Word
(`.docx`) and Excel (`.xlsx`) files and, when saving, rewrites only what you edited.

> 📖 This manual is for people who **use** the application.
> If you are going to work on the code, read the [README](README.md).

> ℹ️ The application speaks Portuguese and English. This manual names the menus and buttons as
> they appear in English; **View → Language** switches the interface.

---

<a id="contents"></a>

## 📑 What is here

| Getting started | Writing | Taking care of the file |
| --- | --- | --- |
| [💿 Install](#install) | [✍️ Text and paragraph](#formatting) | [🔒 Files that open locked](#files-that-open-locked) |
| [📄 Files and templates](#files-it-opens) | [📋 Lists](#lists) | [⚠️ The warnings](#warnings) |
| [⌨️ Shortcuts](#shortcuts) | [📐 Pages, sections and columns](#pages) | [💾 If the application closes on its own](#if-the-application-closes) |
| [👁️ Viewing the document](#view) | [📊 Tables and images](#tables-and-images) | [🛟 Saving does not cost what you did not edit](#saving-does-not-cost) |
| | [🔗 Table of contents and references](#references) | [🖨️ Print and export](#print-and-export) |
| | [📝 Footnotes and endnotes](#notes) | [🚧 Known limits](#known-limits) |
| | [➗ Equations](#equations) | |
| | [💬 Comments and revisions](#review) | |
| | [🔤 Spelling, symbols and counting](#tools) | |
| | [🤖 AI assistant](#ai-assistant) | |
| | [🧮 Spreadsheet formulas](#formulas) | |

---

<a id="install"></a>

## 💿 Install

Nothing here needs internet. **The application does not access the network**, not even to check
spelling. The only exception is the [AI assistant](#ai-assistant), which comes turned off and
only sends text after you configure it and allow it.

### 🐧 Linux — AppImage

Download the file, make it executable and open it:

```bash
chmod +x Librevia-*.AppImage
./Librevia-*.AppImage
```

### 🐧 Linux — `.deb` package

It installs in the applications menu, like any other program:

```bash
sudo apt install ./librevia_*_amd64.deb
```

### 🪟 Windows

Run the installer. It installs **only for your user** and does not ask for an administrator
password. The installer is not signed yet, so Windows may show "Windows protected your PC": click
**More info** and then **Run anyway**. After that, a `.docx`, `.dotx`, `.odt`, `.ott` or `.xlsx`
can be opened with Librevia from the right-click menu in File Explorer.

---

<a id="files-it-opens"></a>

## 📄 The files it opens

| Extension | What it is | What Librevia does with it |
| --- | --- | --- |
| `.docx` | Word document | opens and saves; a new document can also be saved this way |
| `.dotx` | Word template | opening it creates a new document from it; "Save as" also writes templates |
| `.dotm` | Word template with macros | opens like `.dotx`, but the macros are left out, and the application warns |
| `.xlsx` | Excel spreadsheet | opens and saves |
| `.odt` | LibreOffice document | opens and saves; what you did not edit comes back intact |
| `.ott` | LibreOffice template | opening it creates a new document from it; "Save as" also writes templates |
| `.sdoc` | Librevia document | keeps everything, with no loss at all |
| `.ssheet` | Librevia spreadsheet | keeps everything, with no loss at all |
| `.txt` | plain text | opens and saves; saving to it discards the formatting, and the application warns first |
| `.pdf` | PDF document | opens read-only, in a tab of its own, like the others. The status bar shows the page and the zoom; **Edit → Find and replace…** searches the text; a protected PDF asks for the password, which is not stored; **File → Properties…** changes title, author, subject and keywords, and **Save** appends only that at the end of the file; **File → Print…** opens the PDF in the print window, where you can also save a copy |

And these Librevia **creates**. `.pdf` also opens for reading, as the table above shows:

| Extension | How |
| --- | --- |
| `.pdf` | **File → Export to PDF…**, for documents and spreadsheets |
| `.docx`, `.odt`, `.html`, `.md` | **File → Export as**, for documents only |

<a id="tabs"></a>

### Tabs

Each open file has a tab at the top of the window, with an icon for its kind: document,
spreadsheet or PDF. Menus and toolbar act on the file on screen.

- Click a tab to show its file. A dot after the name means unsaved changes.
- **×**, a middle click on the tab, or **File → Close file** closes it. With unsaved changes, the
  application asks first.
- Opening a file that is already open shows its tab, instead of a second copy.
- Up to 16 files stay open at once. Close one to open another.
- The **Home** tab has the recent files and the buttons to create and open. The first file you open
  from it takes its place; **+** brings it back. Closing the last file leaves Home alone.
- A blank document you did not touch gives its place to the next file you open, as in Word. **New
  document** always opens a new tab.
- Hidden tabs keep their drafts too, every eight seconds.

When you come back to a document, the cursor is at the start and **Undo** does not reach what you
did before switching. Spreadsheets keep their undo.

<a id="format-of-new-files"></a>

### The format of new files

The first time you save a new file, Librevia asks which format to use from then on:

| Choice | Documents | Spreadsheets | Who opens them |
| --- | --- | --- | --- |
| **Office and Google** | `.docx` | `.xlsx` | Word, Excel, Google Docs and Sheets, LibreOffice |
| **LibreOffice** | `.odt` | `.xlsx` | LibreOffice and Word; Calc opens the `.xlsx` |
| **Librevia** | `.sdoc` | `.ssheet` | only Librevia, which keeps everything in them |

It asks only once. **Tools → Options…** changes the choice, and **Save as** still offers every format.

<a id="libreoffice-documents-odt"></a>

### LibreOffice documents (`.odt`)

An `.odt` opens and saves like a `.docx`: only the block you edited is written again, and the rest
comes back as it was. "Save as" offers `.odt` for any document: a new document, an `.sdoc` or a
`.docx` go whole into a new `.odt`, which becomes the open file.

These open and come back to the file: footnotes and endnotes, comments with replies, equations,
header and footer (with versions for the first page and for even pages), sections with their own
paper, margins and columns, and tracked changes (inserted text, deleted text and joined
paragraphs). Page number, total pages, date, time, title and author become fields, as in `.docx`.

What Librevia does not show (tracked formatting changes, index markers, sections without columns or
inside another, text boxes, drawings, forms, other fields) stays in the file as long as you do not
edit the block it is in. The application warns on opening; if you edit that block, it is lost
there, and the warning appears on saving.

**`.ods` spreadsheets do not open.** Ask whoever sent it to save it as `.xlsx`. LibreOffice does
this through "Save as".

### Templates

A template is a document ready to serve as a starting point, like letterhead paper.

**File → New from template…** shows the templates that come with the application — blank
document, letter, report with cover and table of contents, and meeting minutes — and also yours,
which live in the templates folder. The **Open templates folder** button shows that folder, and
**Browse…** accepts any `.dotx` or `.ott` on the computer.

The document created is new and untitled. Styles, header, footer and page size come from the
template. When saving, the application asks where to save, and **the template is never
overwritten**.

To create a template of your own, use **Save as → Word template (.dotx)** or **ODT template
(.ott)** and save it in the templates folder.

A document created from an `.ott` gets the `.odt` name. Saved as `.odt`, it writes over the
template's package: what you did not edit comes back as it was in the template.

### Document properties

**File → Properties…** has three parts:

- **Summary**: title, subject, authors, manager, company, category, status (for example, "Draft"
  or "Final"), keywords and comments. You can fill them in and change them.
- **Custom**: fields with their own names, such as "Client" or "Case no.", each of type text,
  number, date or yes or no. **Add property** creates one; the **×** deletes it. They are the same
  as in the **Custom** tab of the advanced properties in Word and Excel. A type Librevia does not
  show, such as a list, stays in the file.
- **Information and statistics**: when the document was created, modified and last printed, by
  whom, the revision number, the total editing time and the language, when the file provides it.

When saving a `.docx` with changed text, Librevia also writes the pages, words, characters and
paragraphs, as Word does. The file manager and the system search read these numbers without
opening the document. If only the properties changed, that part of the file stays as it was.

A spreadsheet works the same way. The statistics count the tabs, the filled cells and the
formulas. The properties go into the `.xlsx`, and Excel shows them in **File → Info**.

To bring a property into the text, use **Insert → Document property…**: choose, for example,
"Client", and the text shows "ACME". When you change the value and click **Apply** in the
properties, the text starts showing the new value. It is the same `DOCPROPERTY` (or `TITLE`,
`AUTHOR`…) field as in Word, which updates it with `F9`. It only works in the body of the text
for now, not in the header or footer.

### Inspect document

Before sending a file out, **File → Inspect document…** shows what it reveals beyond its content,
and removes what you tick:

- **Names of people and of the company**: author, last saved by, manager, company and who wrote
  each comment and change. The authors become "Author", as Word does, and the file starts asking
  that nothing be signed with a name: new comments and changes also come out as "Author", here, in
  Word and in Excel.
- **Computer folders**: the path of the template attached to a Word document, or the last folder
  where Excel saved the spreadsheet. They usually carry the user name, such as
  `C:\Users\ana\...`.
- **Custom properties**: fields created by people or by company systems.
- **First-page thumbnail**: the image the file manager shows.
- **Image metadata**: the camera or program data stored in JPEG and PNG photos, sometimes with the
  GPS location. The image and its orientation stay.

Removing marks the file as changed: only **Save** writes the clean version, and closing without
saving leaves the file as it was. In a `.docx`, `.odt` or `.xlsx` with unsaved changes, the
inspector asks you to save first, because it reads the file as it is saved. In an `.sdoc` or
`.ssheet` there are only names to remove.

In an `.odt`, the inspector looks for the same data. The format has no request not to record
names; Librevia writes its own into the file and respects it, but LibreOffice ignores it.

In a PDF, the inspector only shows: the author in the properties and how many earlier versions
the file keeps. Each change to a PDF, made here or in another program, is appended at the end and
leaves the earlier version inside the file, with the author and title it had. For a PDF without
that data, export it again from the original document.

---

<a id="shortcuts"></a>

## ⌨️ Shortcuts

They are Word's shortcuts, with few exceptions, explained at the end of the table.

### File and editing

| Shortcut | What it does |
| --- | --- |
| `Ctrl+N` | new document |
| `Ctrl+Shift+N` | new spreadsheet |
| `Ctrl+O` | open |
| `Ctrl+S` | save |
| `Ctrl+Shift+S` | save as |
| `Ctrl+W` | close the file |
| `Ctrl+P` | print |
| `Ctrl+F` | find and replace |
| `Ctrl+Shift+V` | paste without formatting |
| `Ctrl+Shift+G` | word count |
| `Ctrl+Shift+Y` | open and close the AI assistant |

### Text and paragraph

| Shortcut | What it does |
| --- | --- |
| `Ctrl+B` / `Ctrl+I` / `Ctrl+U` | bold, italic, underline |
| `Ctrl+Shift+=` / `Ctrl+=` | superscript and subscript (also `Ctrl+.` and `Ctrl+,`) |
| `Ctrl+L` / `Ctrl+E` / `Ctrl+R` / `Ctrl+J` | align left, center, right, justify |
| `Ctrl+1` / `Ctrl+5` / `Ctrl+2` | single, 1.5 and double line spacing |
| `Ctrl+Alt+1` … `Ctrl+Alt+6` | Heading 1 to Heading 6 |
| `Ctrl+Alt+0` | turns the paragraph back into body text |
| `Ctrl+]` / `Ctrl+[` | increase and decrease the indent |
| `Ctrl+Shift+8` / `Ctrl+Shift+7` | bulleted list and numbered list |
| `Tab` / `Shift+Tab` | in a list, moves the item down a level or back up |

### Insert

| Shortcut | What it does |
| --- | --- |
| `Ctrl+Enter` | page break |
| `Ctrl+F12` | table (asks for rows and columns) |
| `Ctrl+Alt+F` | footnote |
| `Ctrl+Alt+D` | endnote |
| `Ctrl+Alt+M` | comment |
| `Alt+=` | equation |
| `Ctrl+Shift+F5` | bookmark |
| `F9` | update fields (page numbers, references) |

### View and review

| Shortcut | What it does |
| --- | --- |
| `Ctrl+F5` | navigation pane |
| `Ctrl+F10` | show and hide formatting marks |
| `Ctrl+F11` | reading mode (`Esc` leaves) |
| `Ctrl+Shift+E` | turn track changes on and off |
| `Ctrl` + numeric keypad `+` / `Ctrl+-` | zoom in and out |
| `Ctrl` + mouse wheel | zoom in and out, one step per notch |
| `Ctrl+0` | back to actual size |

> ℹ️ **The exceptions to Word, and why.**
>
> - **Body text** is `Ctrl+Alt+0`, not `Ctrl+Shift+N`, because here that combination opens a new
>   spreadsheet.
> - **Zoom in** uses the numeric keypad `+`, so as not to take `Ctrl+Shift+=` from superscript.
>   The **View** menu also has zoom in and out. In a PDF, which has no text to format, `Ctrl+=`
>   and `Ctrl++` also zoom in.
> - **Formatting marks** are `Ctrl+F10`, as in LibreOffice. Word's `Ctrl+*` is already used by
>   the bulleted list.

### In the spreadsheet

`Ctrl+B`, `Ctrl+I` and `Ctrl+U` apply to the selected cells. **Right-clicking** a cell opens the
menu to insert and delete rows and columns.

`Ctrl+Z` undoes the last spreadsheet edit (a cell, a paste, an inserted row, a new tab) and
`Ctrl+Y` or `Ctrl+Shift+Z` redoes it. There are up to 100 steps, and the history starts over with
each file opened. While you type in a cell or in the formula bar, `Ctrl+Z` undoes only the typing.

**The fill handle** is the little square at the bottom-right corner of the selection. Drag it, and
the following cells are filled. If the source cell has a formula, it **follows along**: `=B2*C2`
dragged down becomes `=B3*C3`. A reference with `$` does not move.

---

<a id="formatting"></a>

## ✍️ Text and paragraph

The document toolbar has, from left to right: paragraph style, font, size, character formatting,
paragraph, lists, what can be inserted and the page.

### Characters

Besides bold, italic, underline and strikethrough:

| Button | What it does |
| --- | --- |
| **Superscript** and **Subscript** | put the passage above or below the line: `m³`, `H₂O` |
| **All caps** | shows capitals without changing what you typed; clearing the formatting brings the text back as it was |
| **Small caps** | small capitals, the height of the lowercase letters |
| **Font color** | the color of the letters |
| **Text background color** | paints the background of the passage, in any color |
| **Highlight** | Word's highlighter, which has fourteen fixed colors |

> ℹ️ **Why two kinds of background?** In a Word file they are two different things: the
> highlighter and the passage's shading. A document you receive may have both, and the
> application keeps each where it was.

### Styles

A style is formatting with a name. Instead of making each heading bold, 14 pt and blue, one by
one, you apply the **Heading 1** style. If you later decide headings should be green, you change
the style once and all the headings change together.

In a `.docx` made in Word, almost all the formatting lives in the styles. That is why a heading
may show up as a colored band, and not just as large letters.

The **A≡** button, next to the font size, opens the **Styles** pane. It shows the styles **this
document** defines and says which one the paragraph at the cursor uses. There you can:

- **Apply** a style to the paragraph or to the selected passage;
- **Modify…** a style: name, the style it is based on, the style of the next paragraph, font and
  size. **Update to match selection** copies the formatting of the selected text into the style;
- create a **New style…**, for paragraphs or characters;
- **Clear formatting**, which removes direct formatting and leaves only the style's.

The list can show all styles, only paragraph styles or only character styles.

New documents start in **Calibri 11**, as in Word.

### The font list

The selector shows, in this order:

1. **the fonts the open document uses**;
2. the five that come with the application: Calibri, Cambria, Arial, Times New Roman and Courier
   New;
3. **the ones installed on your computer**, in alphabetical order.

The five in the middle work the same on any computer, because they go inside the installer. Yours
work on yours. On a computer that does not have them, the document shows up in another font, and
the pages may break in different places.

### The paragraph dialog

The **¶** button, or **Format → Paragraph…**, gathers on one screen everything that belongs to the
paragraph, already filled in with what it has:

- **alignment**: left, center, right, justified;
- **line spacing**: single, 1.15, 1.5, double, any multiple or a measure in points;
- **spacing** before and after, in points;
- **indentation** left and right, in millimeters;
- **first line**: indent (the first line moves in) or hanging (it moves out, and the others move
  in, as in a bibliographic reference);
- **keep with next**: the paragraph does not stay alone at the foot of the page, away from what
  comes after.

It applies to every paragraph the selection touches. `Enter` applies; `Esc` closes without
applying.

> ℹ️ **"Single" is not "1.0".** Single line spacing is the height the font itself asks for. No
> number reproduces it exactly, and swapping one for the other changes where the pages break.

---

<a id="lists"></a>

## 📋 Lists

`Ctrl+Shift+8` starts a bulleted list; `Ctrl+Shift+7`, a numbered list. The same buttons are on
the toolbar.

A list can have several levels, like an index: `1.`, then `1.1.`, then `1.1.1.`. **Tab** moves
the item down a level, and **Shift+Tab** brings it back. The numbering follows the same rules as
Word, so a document you receive shows the same numbers it showed there.

**List format…** chooses the appearance:

- **List library**: `1. a. i.`, `1. 1.1. 1.1.1.`, `I. A. 1. a.`, `1) a) i)` and three kinds of
  bullets;
- or level by level: the number format (`1, 2, 3`, `01, 02`, `a, b, c`, `A, B, C`, `i, ii`,
  `I, II` or bullet), the text around it and the starting number. The **preview** shows the
  result before applying.

**Right-clicking** a numbered item offers:

| Option | When to use it |
| --- | --- |
| **Restart at 1** | the new list started right after the previous one and inherited its numbering |
| **Continue numbering** | the opposite: a regular paragraph in the middle split the list in two |
| **Set numbering value…** | the list needs to start at another number, such as `5.` |

---

<a id="pages"></a>

## 📐 Pages, sections and columns

### The screen shows the real sheets

The document appears divided into sheets, as it will come out on paper. The break is made **line
by line**, the same way Word and LibreOffice do it:

- a paragraph can start on one sheet and end on the next;
- a line does not stay alone at the top or bottom of the sheet (these are "widows" and
  "orphans");
- a table's header row repeats at the top of each sheet the table continues on, even when a row
  continues there;
- a table row can also start on one sheet and end on the next, unless Word marked it to stay
  whole ("Allow row to break across pages" off). A table inside a cell breaks between its rows.

The page count in the status bar is the same as in the PDF and on paper.

### Page setup

**File → Page setup…** defines:

- paper **size** and **orientation** (portrait or landscape);
- **margins**, in millimeters;
- **header and footer**, with the **Page number** and **Total pages** buttons;
- **Different first page**, for a cover without a number;
- **Different odd and even pages**, for a book where the number always sits on the outer side;
- **page numbering**: the format (`1, 2, 3`, `i, ii, iii`, `A, B, C`…) and the number it starts
  at.

Header and footer can also be edited directly on the sheet: click the text and type.

### Sections

A section is a part of the document with its own page setup. It is useful, for example, to put a
wide table on a landscape sheet in the middle of a portrait report, or to number the introduction
with `i, ii, iii` and the rest with `1, 2, 3`.

**Insert → Section break** offers four kinds:

| Kind | The new section starts… |
| --- | --- |
| **Next page** | on the next sheet |
| **Continuous** | on the same sheet, right below |
| **Even page** | on the next even sheet |
| **Odd page** | on the next odd sheet, like a book chapter |

With sections, page setup asks where to **Apply to**: **This section** or **Whole document**. A
section's footer can **link to the previous one**, that is, repeat the one of the section above,
or have its own text.

**Insert → Delete section break** joins the section with the one above.

### Columns

**Format → Columns…** divides the section into columns, as in a newspaper: the **number of
columns**, the **space between them**, in millimeters, and a **line between columns**, if you
want.

**Insert → Column break** sends the following text to the top of the next column.

---

<a id="tables-and-images"></a>

## 📊 Tables and images

**Insert table** is in the **Table** menu, on the toolbar's table button and at `Ctrl+F12`, and
asks how many rows and columns you want. The first row comes marked as the **header**: it is the
row that repeats at the top of each sheet when the table continues on the next.

With the cursor in a cell, the **Table** menu and the **right-click** menu offer:

- insert a row above or below, a column to the left or right;
- delete a row, a column or the whole table;
- **Merge cells** (select them by dragging the mouse) and **Split cell**;
- turn the **Header row** on and off;
- **Table properties…**: the width of the cursor's column, in millimeters; the border of the
  selected cells (style, width, color and which sides); the shading; and repeating the header.

The divider between two columns can also be **dragged** with the mouse, at any zoom.

> ℹ️ The dialog only offers what `.docx` stores. Borders Word has and the screen does not draw,
> such as thick-and-thin double or wavy lines, show up as a single line, and **come back intact**
> on saving, as long as you do not format that cell. If you do, the loss warning says so.

**Merging**: cells merged in a `.docx`, horizontally or vertically, stay merged, and the ones you
merge here also go into the file.

**Resizing an image**: click it and drag one of the eight handles. At the corners, the proportion
is locked; hold `Shift` to stretch freely. The side handles change only one measure. The image
does not go beyond the width of the text column, and one `Ctrl+Z` undoes the whole drag. With a
handle selected, the arrow keys also resize.

**Alt text and alignment**: with the image selected, the toolbar's image button becomes **Image
properties** (also in **Format → Image…**). Alt text is what a screen reader reads in place of the
image, and it goes into the file.

---

<a id="references"></a>

## 🔗 Table of contents and references

### Navigation pane

**View → Navigation pane** (`Ctrl+F5`) shows the document's headings in a column at the side, like
an index. Click a heading to go to it. In a hundred-page document, it is the fastest way to find a
chapter.

### Table of contents

**References → Insert table of contents** builds the table of contents from the paragraphs with the
**Heading 1**, **Heading 2** styles and so on, with each one's page number.

The table of contents does not change on its own as you write. After adding chapters or after the
pages move, use **References → Update table of contents**.

### Bookmarks

A bookmark is a name given to a point in the document, like a bookmark in a book. **Insert →
Bookmark…** (`Ctrl+Shift+F5`) creates, lists, goes to or deletes a bookmark. The list can be
sorted by name or by position in the document.

The **link** dialog also points inside the document: instead of a web address, choose a heading or
a bookmark.

### Captions

**References → Insert caption…** numbers figures, tables and equations: `Figura 1`, `Figura 2`…
Choose the label, the text after the number and whether the caption goes above or below the item
where the cursor is.

### Cross-reference

A cross-reference is a phrase such as "see Figure 3, on page 12" where the **3** and the **12**
update on their own. **References → Cross-reference…** asks:

- the **type**: heading, bookmark, footnote, endnote or caption;
- **which one**;
- what to insert: the **text**, the **number**, the **note number** or the **page number**;
- and whether the reference should work as a **hyperlink**, going to that point when clicked.

### Update fields

Page numbers, total pages, cross-references, captions and document properties are **fields**:
text the application calculates. `F9`, or **References → Update fields**, recalculates them all.

---

<a id="notes"></a>

## 📝 Footnotes and endnotes

A **footnote** sits at the foot of the same sheet; an **endnote** sits at the end of the document.

- **Insert → Footnote** (`Ctrl+Alt+F`)
- **Insert → Endnote** (`Ctrl+Alt+D`)

The number appears in the text, and the cursor goes to the note, **at the foot of its own sheet**,
where you write. The note numbers follow the order in which they appear in the text.

**Right-clicking** a note lets you **convert** it from footnote to endnote, or the other way
around.

---

<a id="equations"></a>

## ➗ Equations

The equations in a Word document appear drawn on the sheet. New ones are written in **LaTeX**, the
way of writing mathematics used in scientific papers: `\frac{a}{b}` is a fraction, `x^2` is a
square, `\sqrt{x}` is a root.

- **Insert → Equation** (`Alt+=`) puts the equation in the middle of the line, with the text.
- **Insert → Display equation** puts the equation on a line of its own, centered.

The dialog has the **LaTeX** field, a **Preview** that follows what you type and a palette of
**Templates** for those who do not know the commands by heart: fractions, roots, indices, sums,
integrals, accents, functions such as sine and logarithm, and symbols such as ∈, ∀ and ⇒.

To change an equation, **double-click** it.

> ℹ️ A Word equation with some construct the screen cannot draw appears **locked**, with what
> could be drawn. It goes back to the file whole, exactly as it came.

---

<a id="review"></a>

## 💬 Comments and revisions

### Comments

Select the passage and use **Insert → Comment** (`Ctrl+Alt+M`). The comment appears in the pane at
the side, linked to the passage. **View → Comments** shows and hides that pane.

On each comment you can **Reply**, **Edit**, **Resolve** (it is marked as resolved, but stays in
the file), **Reopen** and **Delete**. **Insert → Next comment** and **Previous comment** move from
one to the next.

The name that appears on your comments and revisions is set in **Tools → Author name…** or in **Tools → Options…**.

Comments in a `.docx` come along and go back to the file, replies included.

### Track changes

With track changes on, each change is recorded instead of simply made: inserted text appears
underlined, deleted text appears struck through, each with the name of whoever changed it.
Whoever receives the document decides what stays.

**Review → Track changes** (`Ctrl+Shift+E`) turns it on and off. The **Review** menu also has:

- **Accept change** and **Reject change**, for the one at the cursor;
- **Accept all changes** and reject all;
- **Next change** and **Previous change**.

**Review → Show** chooses how to see the document:

| Mode | What appears |
| --- | --- |
| **All markup** | everything in view: the inserted underlined and the deleted struck through |
| **Simple markup** | the text as it will be, with a bar in the margin next to each changed paragraph |
| **No markup** | the text as it will be, clean |
| **Original** | the text from before the changes; in this mode the editor is read-only |

Switching modes does not change the document, and printing follows the chosen mode.

---

<a id="tools"></a>

## 🔤 Spelling, symbols and counting

### Spelling is in Portuguese and does not use the internet

The Brazilian Portuguese dictionary **comes inside the installer**. The first time the application
opens, it is copied to your user folder, and that is where the checker looks for it.

Why does this matter? The checker that comes with Chromium, on its own, would download the
dictionary from the internet. On a computer without a network, it would simply mark nothing, and
it would not warn you.

A misspelled word gets the usual wavy underline, **in the text and also in the header and
footer**, where a typo repeats on every sheet.

**Right-click** the underlined word to:

- choose one of the suggestions, which replaces the word;
- **Add to dictionary**, so you are never warned about it again, in any document;
- **Ignore in this session**, which lasts until you close the application. It is useful for the
  name of a client that appears ten times in this document and in no other.

To turn it off: **Tools → Spell check**.

### Right-clicking the document

Besides the spelling suggestions, the menu has **cut**, **copy**, **paste** and **paste without
formatting**. Inside a table, it also brings the table's actions. A greyed-out item does not apply
there: "paste" is greyed out when nothing has been copied, for example.

### Paste without formatting

`Ctrl+Shift+V`, the **Edit** menu or the right-click menu. The text comes in with the document's
formatting, not with the formatting of where it came from. It is what keeps a passage copied from
a website from arriving with another font, another size and another color. Each line break in the
copied text becomes a paragraph.

### Word count

`Ctrl+Shift+G` or **Tools → Word count…** shows words, characters with and without spaces,
paragraphs and pages, in two columns: the **document** and the **selection**. With nothing
selected, the selection column shows a dash, because "nothing selected" is not the same as "zero
words". The dialog follows what you type, so it can stay open.

An empty paragraph does not count: it is an `Enter` pressed to open space, and Word does not count
it either.

### Special characters

**Insert → Special character…** opens a pane with what the keyboard usually lacks: typographic
quotes, dashes, currency symbols, math signs, Greek letters and marks such as © and ®. Click the
symbol, or move around the grid with the arrows and insert with `Enter`. The pane stays open,
because whoever opens it usually wants more than one symbol.

The `␣` in the "Punctuation" group is the **non-breaking space**. It is what keeps `R$` from
sitting at the end of one line and the amount at the start of the next.

### Formatting marks (¶)

`Ctrl+F10`, the toolbar's `¶` button or **View → Formatting marks** show what exists in the text
but has no ink: a dot for each space, an arrow for each tab, a `¬` for each line break and a `¶`
at the end of each paragraph. They help find out why an alignment came out wrong. It is almost
always an extra space, or a tab where there should be an indent.

The marks **do not come out on paper or in the PDF**, and they do not change where the pages
break.

### Smart typography

It comes turned on. As you type, it replaces:

| What you type | What comes out |
| --- | --- |
| `"quotes"` | `“quotes”` |
| `'quotes'` | `‘quotes’` |
| `--` | `—` (em dash) |
| `...` | `…` |
| `(c)` `(r)` `(tm)` | `©` `®` `™` |
| `1/2` `1/4` `3/4` | `½` `¼` `¾` |
| `+/-` `!=` `<<` `>>` | `±` `≠` `«` `»` |

It is what Word does in Portuguese. `->` and `3 x 4` do **not** become an arrow and `×`, on
purpose: in technical text that gets in the way more than it helps.

A `Backspace` right after a replacement undoes **only that replacement**. Write `--silent`, delete
once, and the dash goes back to being two hyphens. It is the way to write a command without
turning anything off.

To turn it off: **Tools → Smart typography**. What has already been replaced stays replaced.

---

<a id="ai-assistant"></a>

## 🤖 AI assistant

The assistant rewrites, summarizes, proofreads or translates the passage you select, and chats
about the document. In the spreadsheet, it explains and writes formulas and fills columns (see
[In spreadsheets](#in-spreadsheets)).

Two guarantees hold throughout this chapter:

- it comes **turned off**: as long as no assistant is configured, nothing leaves the computer;
- **nothing changes in the document or the spreadsheet without a click from you.**

### Configure

Open **Tools → AI assistants…**. There are two ways:

- **Find assistants** looks, only when you click, at what is already on this computer: Ollama
  (`localhost:11434`), LM Studio (`localhost:1234`), the `claude` program (Claude Code) and the
  installed ACP agents (see below). Whatever answers appears with an **Add** button.
- **Add manually** accepts an OpenAI-compatible server (address and, if it asks, a key), the
  Anthropic API (key) or an ACP agent.

**ACP agents.** ACP (Agent Client Protocol) is an open protocol that several AI agents speak. The
search finds these, if the program is on the `PATH`: Gemini CLI (`gemini`), Claude Agent
(`claude-agent-acp`), Codex (`codex-acp`), OpenCode, Goose, GitHub Copilot (`copilot`), Qwen Code
(`qwen`), Kimi CLI, Cursor Agent (`cursor-agent`), Kilo Code (`kilo`), Auggie, Cline, Mistral Vibe
(`vibe-acp`) and Amp (`amp-acp`). An agent installed elsewhere comes in through **Add manually →
ACP agent → Choose the program…**, with the arguments it asks for.

- Login happens in the agent itself: open it once in the terminal and sign in to your account.
- Librevia does not download or install agents, does not offer them your files or a terminal, and
  runs them in an empty temporary folder.
- The agent's own tools (those of Gemini CLI, Codex, Claude Agent…) follow its configuration, not
  Librevia's. To restrict what it can do, configure the agent.
- On Windows, only agents that are an `.exe` work; the ones installed by npm (`.cmd`) do not.

The key is encrypted by the system's password vault and never comes back to the screen. On Linux
without a vault (no GNOME Keyring or KWallet), Librevia refuses to store the key; local assistants
keep working, because they do not need one.

In the same dialog you choose the default assistant. **Edit** changes an assistant's name, address
and key; a blank key keeps the stored one. Changing the address erases the key and the
permission, so they do not follow along to another server. On an OpenAI-compatible server,
**Maximum context** says how many characters of the document go along (see
[What is sent](#what-is-sent-and-where)).

In **Preferences**, in the same dialog:

- **Keep the formatting in revisions** (on by default): when rewriting, proofreading or
  translating, the passage goes with bold, italic, strikethrough, headings, lists and links, and
  the answer comes back with them. Turned off, the passage goes and comes back as plain text.
- **If the Anthropic model refuses, try another model Anthropic recommends** (off by default): it
  only applies to the Anthropic API. If there is a switch, the author of the change is the model
  that finished the answer.

### Use

1. Select the passage.
2. Right-click and open **AI assistant ▸**: Rewrite, Summarize, Fix spelling and grammar,
   Translate or Custom request. From the keyboard, `Shift+F10` opens the menu. The panel also
   opens through **Tools → AI assistant…** (`Ctrl+Shift+Y`), and the same shortcut closes it.
3. In the **Actions** tab, **Model** chooses the model: **Automatic** uses the assistant's
   default, and **List models** asks which ones it has.
4. The answer appears in the panel as it is written, already formatted. **Stop** interrupts it.
5. Choose what to do with it:
   - **Apply** replaces the passage with the answer as a **tracked change**, with the model's name
     as author, even with track changes turned off;
   - **Insert below** puts the answer in new paragraphs after the passage, also marked;
   - **Copy** takes the answer to the clipboard.

If you edit the passage while the answer arrives, **Apply** is greyed out and the panel warns that
the passage changed. Insert below and Copy still work.

### Chat

The **Chat** tab exchanges several messages with the assistant about the open document. `Enter`
sends; `Shift+Enter` breaks the line. Each answer has **Insert below**, which puts the text after
the selection or the cursor as a tracked change, and **Copy**. **New chat** starts from scratch.

With each message goes the document **as it is at that moment**. With a passage selected, the
model understands "this" as that passage.

The chat lives only in memory: it does not go into the file or onto the disk, and it disappears
when you close the document.

#### Proposed edits and comments

Below the message field is the **Let the assistant propose edits and comments** box. It comes
turned off. Turned on, the assistant can read and search the document and **propose**: replacing
or removing a passage, new text after a paragraph (that is how it writes in an empty document) and
comments. The proposed text comes formatted, with headings, lists, quotes, bold, italic and links.
Tables and images cannot be proposed yet.

Each proposal shows up in two ways, with the **Accept** and **Reject** buttons:

- **on the sheet, as a difference**: the outgoing word struck through in red, the incoming one in
  green, a comment marked in yellow. You can keep writing, and the difference follows the text;
- **in the panel**, below the answer, with **Accept all** and **Reject all**.

Accepting puts the change in the text, and `Ctrl+Z` takes it out. With track changes on, it goes
in as a tracked change, with the model's name as author. Before accepting, Librevia checks that the
paragraph still has the text the model read; if you changed it, Accept says the text changed and
touches nothing.

All assistants offer this. A local model without tool support refuses the request, and Librevia
tells you to turn the box off or switch models.

<details>
<summary><b>The limits of proposals</b></summary>

- At most 30 proposals per message, 10 tools per round and 8 rounds.
- What the model reads from the document in one answer does not go beyond the assistant's context
  limit.
- The assistant does not save, does not open files, does not access the internet and does not run
  programs: it only reads the text sent with the message and proposes.
- Claude Code receives only Librevia's five tools, through a key valid for one message, and stays
  without its own tools and without any MCP servers you configured in it.
- An ACP agent receives the same five tools. Librevia only permits their use, but the agent may
  have tools of its own that do not ask for permission.

</details>

<a id="in-spreadsheets"></a>

### In spreadsheets

In the spreadsheet editor, **Tools → AI assistant…** (`Ctrl+Shift+Y`) opens the panel. The
assistant and the model are the same as in the document. There are three tasks:

- **Explain formula** explains the formula in the active cell. Nothing is written.
- **Write a formula**: write what you want (for example, "sum of the prices in column B") and
  click **Generate**. The panel shows the formula and the value it would give. **Insert into
  cell** writes it, and `Ctrl+Z` undoes it. The formula must pass the same check as the formula
  bar.
- **Fill a column**: select the cells, in a single column and with at most 500 rows, and write the
  request (for example, "category of the product in column A"). The panel shows a preview with the
  current value and the new one. **Apply** writes them all, and `Ctrl+Z` undoes them all. Values
  go in, never formulas.

In the spreadsheet there is no chat, and each request is independent.

### External programs (MCP)

Claude Code in the terminal, Claude Desktop or another program that speaks the MCP protocol can
work on the document or spreadsheet open in Librevia, including what has not been saved yet.

1. Open **Tools → AI assistants…** and turn on **Let external programs read the open document and
   propose changes (MCP)**. It comes turned off.
2. Register Librevia in the program, once:
   - **Claude Code:** copy the command shown below it and run it in the terminal. It starts with
     `claude mcp add --scope user librevia --`.
   - **Claude Desktop and others:** copy the JSON configuration and paste it into
     `claude_desktop_config.json` or into the project's `.mcp.json`.
3. With Librevia open, ask the program, for example: "read the document open in Librevia and
   comment on the sentences that are too long".

In the document, the program reads, searches, proposes replacements, inserts text and comments.
**Each proposal goes in right away as a tracked change or a comment**, with the program's name
followed by "(MCP)" as author, for example "Claude Code (MCP)". You accept or reject it like any
revision.

In the spreadsheet, which has no track changes, the program reads the cells and proposes writes,
which **do not go in right away**: they appear in a banner above the grid, with a preview.
**Apply** writes, and `Ctrl+Z` undoes; **Discard** throws it away.

The program does not open or save files, does not accept or reject revisions, does not see the
file's name or path and does not touch PDFs.

- While a program is connected, the status bar shows its name and the **Disconnect** button. At
  most 2 programs at the same time.
- The connection does not use the network: the program runs `librevia --mcp-bridge`, which talks
  to the open Librevia through a communication file (socket) only your user can open, with a key
  that changes at every start.
- On Linux, the program must run inside the graphical session (with `DISPLAY` or Wayland).

<a id="what-is-sent-and-where"></a>

### What is sent, and where

| When | What goes |
| --- | --- |
| **Actions** | the passage, up to 20,000 characters; the instruction of the custom request; and the document's text before and after the passage, as context, without formatting and without what was deleted |
| **Chat** | with each message, the current document, your message (up to 2,000 characters) and the previous messages, up to 40. The previous ones take at most half of the context limit; beyond that, the oldest are left out |
| **Chat with proposals** | the document in numbered paragraphs, without what was deleted or the footnotes. If it does not fit, the beginning goes, and the model asks for the rest |
| **Spreadsheet** | the active sheet: the active cell, the headers and the nearest filled cells, up to 200 characters each. To fill, the other columns of each selected row |
| **MCP program** | the whole text of the document, or the cells it asks for. It sends them wherever it is configured to: Claude Code and Claude Desktop send them to Anthropic |

- **Never sent**: the file name, the properties or the other tabs of the spreadsheet.
- **The context limit** is 200,000 characters, or 6,000 for an assistant on this computer, which
  reads little at a time. In a larger document, the text nearest the passage goes, and the model is
  told that it sees only a part.
- **The model is told** that the document is reference material and that instructions written in
  it do not come from you.
- **An assistant at `localhost` or `127.0.0.1`** runs on this computer: the text does not leave
  it.
- **An assistant outside the computer** (another address, the Anthropic API, Claude Code or an ACP
  agent) asks before the first send, showing the destination. For an ACP agent, the destination is
  the program with its arguments, and changing them brings the question back. **Ask permission
  again** does too.
- Claude Code and ACP agents do not keep the chat, so they receive the whole chat with each
  message. Claude Code reports the cost of the request, and the panel shows it.

### Limits

- With **Keep the formatting** on, underline, color and font do not pass through the model: the
  answer gets those of the start of the passage.
- A list in the answer uses Librevia's default numbering, not that of a Word list that was in the
  passage.
- The selection cannot have images, equations, notes or fields, nor go from one table cell to
  another.
- With the file read-only, you can ask and chat, but Apply, Insert and Accept are greyed out, and
  proposals from MCP programs are refused.

---

<a id="formulas"></a>

## 🧮 Spreadsheet formulas

```excel
=SOMA(1,5;2)
```

This adds one and a half to two, and gives 3,5. Two rules go together, and they are the same as in
Excel in Portuguese:

- **the comma separates the decimals**;
- **the semicolon separates the arguments**, precisely because the comma is already taken.

If you type `=SOMA(A1,B1)`, the application tells you which separator is right, instead of
showing some error.

The names work in **both languages**: `SOMA` and `SUM` are the same function, as are `SE` and
`IF`, `PROCV` and `VLOOKUP`. A formula copied from a spreadsheet in English does not need to be
translated by hand.

<a id="available-functions"></a>

### Available functions

There are 52, in six groups. Each also answers to its English name (`SOMA` is `SUM`):

| Group | Functions |
| --- | --- |
| **Math** | `SOMA` `SOMASE` `ARRED` `ARREDONDAR.PARA.CIMA` `ARREDONDAR.PARA.BAIXO` `ABS` `INT` `TRUNCAR` `RESTO` `RAIZ` `POTÊNCIA` |
| **Statistics** | `MÉDIA` `MÁXIMO` `MÍNIMO` `CONT.NÚM` `CONT.VALORES` `CONTAR.VAZIO` `CONT.SE` |
| **Logic** | `SE` `SEERRO` `SENÃODISP` `E` `OU` `NÃO` `ÉERROS` `É.NÃO.DISP` `ÉNÚM` `ÉTEXTO` `ÉCÉL.VAZIA` |
| **Text** | `CONCATENAR` `NÚM.CARACT` `ESQUERDA` `DIREITA` `EXT.TEXTO` `MAIÚSCULA` `MINÚSCULA` `ARRUMAR` `SUBSTITUIR` `PROCURAR` `LOCALIZAR` `VALOR` |
| **Lookup** | `PROCV` `PROCH` `CORRESP` `ÍNDICE` |
| **Date** | `HOJE` `AGORA` `DATA` `ANO` `MÊS` `DIA` `DIA.DA.SEMANA` |

### What does not exist yet

Dynamic arrays, whole-column references (`A:A`) and named ranges.

> ✅ A formula with a function outside the list shows `#NOME?` in the cell, but **stays in the
> file, intact**, and Excel calculates it normally again. The warning on opening says which
> functions those are.

---

<a id="view"></a>

## 👁️ Viewing the document

| In the View menu | What it does |
| --- | --- |
| **Zoom in**, **Zoom out**, **Actual size** | zoom from 50 % to 200 %; the status bar buttons do the same |
| **Fit to width** | the sheet takes up the width of the window |
| **Reading mode** (`Ctrl+F11`) | hides the bars and leaves only the text; the document is read-only while it lasts, and `Esc` leaves |
| **Theme** | **Light**, **Dark** or **System**, which follows what the computer uses |
| **Language** | **Português** or **English**, for menus, buttons and messages |
| **Show toolbar**, **Show status bar** | hide the bars to gain space |
| **Full screen** | takes up the whole screen |

Zoom only changes the size on screen. The pages break in the same places at any zoom.

---

<a id="files-that-open-locked"></a>

## 🔒 Some files open locked — and how to unlock them

Some documents open **read-only**, with an orange banner at the top saying exactly what they have.
It happens when the file has something the editor shows but cannot write back if you edit the
passage:

- a **field** of a kind the editor cannot recalculate. The text Word stored appears, but
  rewriting the paragraph would turn it into plain text;
- **content controls**, Word's little form boxes;
- **structural revisions**, such as a cell inserted or deleted in a table with track changes on.

Whoever only needs to read runs no risk at all. Whoever needs to edit clicks **Edit anyway** and
goes on, knowing what the risk is.

It is not a padlock: one click unlocks it, and it applies only to that file. A regular file opens
editable, even if it has comments, revisions or notes, because Librevia knows how to write them
back. Locking every file would teach you to click without reading, and the protection would stop
protecting.

---

<a id="warnings"></a>

## ⚠️ The warnings: read them, they differ from one another

When opening a file, the application may show a warning. There are three kinds, because each one
calls for a different attitude:

| Warning | What it means | What you can do |
| --- | --- | --- |
| **Invisible** | the feature stays in the file, but does not appear here: charts, a shape's shadow or gradient, conditional formatting and spreadsheet filters | edit and save freely: it comes back intact |
| **Loss** | something really vanishes on saving | it only happens if you edit exactly the passage where it is; the warning says which |
| **Lock** | it is the loss that made the file open [read-only](#files-that-open-locked) | read freely; to edit, click **Edit anyway** |

If every warning said the same thing, you would learn to close them without reading.

---

<a id="if-the-application-closes"></a>

## 💾 If the application closes on its own

Every eight seconds, what is on screen is kept in a draft, **never over your file**. If the
computer shuts down or the application closes suddenly, on the next start a blue banner offers the
work back:

- **Recover** brings the content back in a tab of its own, marked as *unsaved*, because that is
  what it is. Check it and save it wherever you want.
- **Discard** deletes the draft.

While the banner is on screen, the draft **is not overwritten**. You can ignore it for a while,
open another file, and the work stays kept.

When you save, the draft is deleted: the file on disk is already the good version.

If only the window stops working, the application asks whether you want to **Reopen**. On
reopening, the same banner offers the draft.

### The error log

When something fails, Librevia writes down the cause in a file, `librevia.log`. It stays **only on
this computer**: nothing is sent. **Help → Open log folder** shows the folder:

| System | Folder |
| --- | --- |
| Linux | `~/.config/Librevia/logs` |
| Windows | `%APPDATA%\Librevia\logs` |

When reporting a problem, attach that file. It carries the version, the system and the error
messages, and it may carry the name of an open file. Past 1 MB, it becomes `librevia.old.log` and
a new one starts.

---

<a id="saving-does-not-cost"></a>

## 🛟 Saving does not cost what you did not edit

Many editors, when saving a `.docx`, write the whole file again, and along the way lose what they
did not understand. Librevia does it differently: **it rewrites only what you touched** and hands
back the rest exactly as it was.

In practice: font, alignment, borders, charts, pivot tables, comments and filters stay in the file
after you fix a comma. An `.odt` works the same way (see
[LibreOffice documents](#libreoffice-documents-odt)).

**New document as `.docx`.** Choose "Office and Google" on the first save (see
[The format of new files](#format-of-new-files)), or "Word document" in "Save as". Since there is no source file,
the application creates a minimal package, with styles, page and properties, but without your name
inside, and writes the document into it. Each save starts from the same minimal package, not from
the file saved last time. So saving ten times gives the same file as saving once.

---

<a id="print-and-export"></a>

## 🖨️ Print and export

### Paper and PDF

What comes out on paper is **what is on screen**. The PDF is generated by the same engine that
draws the editor, with selectable text and embedded fonts.

- **Documents** come out with the header, footer, numbering, margins, orientation and columns of
  each section.
- **Spreadsheets** come out with the whole active tab, fitted to the page width. Frozen rows
  become a header and repeat on each sheet.

**File → Print preview** shows the sheets before spending paper.

> 💡 If the spreadsheet columns are cramped, put the page in landscape in **File → Page
> setup…**.

### Export as Word, ODT, HTML or Markdown

In **File → Export as**, the document becomes a **web page** (`.html`), a **Markdown** text
(`.md`) or an **OpenDocument text** (`.odt`). It is a new file: the document stays open in the same
place, as it was.

**Word document (.docx)** is different: it works like **Save as**, only with Word's format. The
document becomes the `.docx`, and the next **Save** writes to it.

- **HTML** comes out on a single page, with the styles, lists, tables and images inside it. The
  notes go to the end, with a link back to the text.
- **Markdown** carries headings, emphasis, links, lists, tables and notes (`[^1]`). The images go
  to a folder next to it, `name_files/`. A table with merged cells comes out in HTML, because
  Markdown has no way to merge.
- **ODT** opens in LibreOffice and in other editors that read OpenDocument, with the styles,
  numbered lists, tables, images, notes, bookmarks and links. Each section carries its own paper,
  margins, columns, header and footer, and the comments go as annotations. To work on the `.odt`
  itself, use **Save as**.
- In ODT, the **revisions** go as tracked changes: inserted text, deleted text and joined
  paragraphs, with author and date. LibreOffice shows them for you to accept or reject. In HTML and
  Markdown, the revisions come out already accepted and the **comments** are left out.
- **Equations** go along in all formats.

---

<a id="known-limits"></a>

## 🚧 Known limits

| Limit | What happens |
| --- | --- |
| **Files over 20 MB** | do not open |
| **`.ods`** | does not open |
| **Macros** | a `.dotm` opens, but without the macros |
| **Merging cells in the spreadsheet** | does not exist yet |
| **Spreadsheet filters** | are preserved in the file, but there is no screen to create or change them |
| **Formulas** | no dynamic arrays, whole-column references (`A:A`) or named ranges |
| **Comments** | the comment text is plain text, without bold or color |
| **Track changes** | **formatting** changes (making something bold, changing the font) are not recorded as revisions |
| **Equations** | use the computer's math font; some LaTeX spacing commands, such as `\,` and `\quad`, are lost when saving as `.docx` |
| **Table of contents** | **Update table of contents** redoes all the entries, including the ones you edited by hand |
