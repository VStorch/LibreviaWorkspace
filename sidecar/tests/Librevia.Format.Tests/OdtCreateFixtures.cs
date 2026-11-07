namespace Librevia.Format.Tests;

/// <summary>
/// A model with everything the ODT writer handles, as the editor sends it: the editor's built-in
/// styles, a plain header with page fields, a section, comments, notes, a list with a start, a
/// merged table and an image.
/// </summary>
public static class OdtCreateFixtures
{
    public const string RichModel = """
        {
          "page": {
            "size": "A4",
            "orientation": "portrait",
            "margins": {"top": 25, "right": 25, "bottom": 25, "left": 25},
            "header": "Relatório — página {n} de {total}",
            "footer": "",
            "headerBand": null,
            "footerBand": null,
            "firstHeaderBand": null,
            "firstFooterBand": null,
            "evenHeaderBand": null,
            "evenFooterBand": null,
            "headerDistanceMm": 12.5,
            "footerDistanceMm": 12.5
          },
          "sections": [
            {
              "size": "A4",
              "orientation": "portrait",
              "margins": {"top": 30, "right": 20, "bottom": 20, "left": 30},
              "header": "",
              "footer": "",
              "headerBand": null,
              "footerBand": null,
              "firstHeaderBand": null,
              "firstFooterBand": null,
              "evenHeaderBand": null,
              "evenFooterBand": null,
              "headerDistanceMm": 12.5,
              "footerDistanceMm": 12.5,
              "id": "s1"
            }
          ],
          "properties": {
            "title": "Relatório <anual> & \"final\"",
            "subject": "Contas",
            "creator": "Ana",
            "keywords": "contas; balanço",
            "created": "2026-01-02T03:04:05Z",
            "modified": "2026-02-03T04:05:06Z"
          },
          "comments": [
            {
              "id": "9",
              "author": "Bia <revisora>",
              "date": "2026-03-01T10:00:00Z",
              "paragraphs": ["Confira & ajuste"],
              "done": false
            },
            {
              "id": "10",
              "parentId": "9",
              "author": "Ana",
              "date": "",
              "paragraphs": ["Feito"],
              "done": false
            }
          ],
          "doc": {
            "type": "doc",
            "content": [
              {
                "type": "heading",
                "attrs": {"level": 1},
                "content": [
                  {"type": "bookmarkStart", "attrs": {"name": "_Toc1", "bid": "1"}},
                  {"type": "text", "text": "Título <script>&"},
                  {"type": "bookmarkEnd", "attrs": {"bid": "1"}}
                ]
              },
              {
                "type": "tableOfContents",
                "attrs": {"instr": "TOC \\o \"1-3\" \\h", "head": 0},
                "content": [
                  {
                    "type": "paragraph",
                    "attrs": {"styleId": "TOC1"},
                    "content": [
                      {
                        "type": "text",
                        "text": "Introdução",
                        "marks": [{"type": "link", "attrs": {"href": "#_Toc1"}}]
                      },
                      {"type": "text", "text": "\t"},
                      {"type": "field", "attrs": {"instr": "PAGEREF _Toc1 \\h", "result": "1"}}
                    ]
                  }
                ]
              },
              {
                "type": "paragraph",
                "attrs": {
                  "textAlign": "center",
                  "spaceBefore": 6,
                  "spaceAfter": 12,
                  "lineHeight": "1.5",
                  "keepNext": true
                },
                "content": [
                  {"type": "text", "text": "Olá  "},
                  {"type": "text", "text": "negrito", "marks": [{"type": "bold"}]},
                  {"type": "text", "text": " "},
                  {"type": "text", "text": "itálico", "marks": [{"type": "italic"}]},
                  {"type": "text", "text": " sub", "marks": [{"type": "underline"}, {"type": "strike"}]},
                  {
                    "type": "text",
                    "text": " cor",
                    "marks": [
                      {
                        "type": "textStyle",
                        "attrs": {"color": "#c00000", "fontFamily": "Arial, sans-serif", "fontSize": "14pt"}
                      }
                    ]
                  },
                  {
                    "type": "text",
                    "text": " realce",
                    "marks": [{"type": "highlight", "attrs": {"color": "yellow"}}]
                  },
                  {
                    "type": "text",
                    "text": " link",
                    "marks": [{"type": "link", "attrs": {"href": "https://example.com/?a=1&b=2"}}]
                  },
                  {
                    "type": "text",
                    "text": " perigoso",
                    "marks": [{"type": "link", "attrs": {"href": "javascript:alert(1)"}}]
                  },
                  {
                    "type": "noteRef",
                    "attrs": {"kind": "footnote"},
                    "content": [{"type": "paragraph", "content": [{"type": "text", "text": "Nota de rodapé"}]}]
                  },
                  {"type": "commentStart", "attrs": {"cid": "9"}},
                  {"type": "text", "text": " comentado"},
                  {"type": "commentEnd", "attrs": {"cid": "9"}},
                  {"type": "text", "text": " removido", "marks": [{"type": "deletion"}]},
                  {"type": "text", "text": " inserido", "marks": [{"type": "insertion"}]},
                  {"type": "hardBreak"},
                  {"type": "field", "attrs": {"instr": "PAGE", "result": "1"}},
                  {
                    "type": "image",
                    "attrs": {
                      "src": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAADAAAAAAAAAAAA",
                      "alt": "figura",
                      "width": 96,
                      "height": 48
                    }
                  }
                ]
              },
              {
                "type": "orderedList",
                "attrs": {
                  "numbering": {
                    "key": "k1",
                    "levels": [
                      {"fmt": "decimal", "text": "%1.", "start": 3},
                      {"fmt": "lowerLetter", "text": "%1.%2)", "start": 1}
                    ]
                  }
                },
                "content": [
                  {
                    "type": "listItem",
                    "content": [{"type": "paragraph", "content": [{"type": "text", "text": "três"}]}]
                  },
                  {
                    "type": "listItem",
                    "content": [
                      {"type": "paragraph", "content": [{"type": "text", "text": "quatro"}]},
                      {
                        "type": "orderedList",
                        "content": [
                          {
                            "type": "listItem",
                            "content": [{"type": "paragraph", "content": [{"type": "text", "text": "sub"}]}]
                          }
                        ]
                      }
                    ]
                  }
                ]
              },
              {
                "type": "bulletList",
                "content": [
                  {
                    "type": "listItem",
                    "content": [{"type": "paragraph", "content": [{"type": "text", "text": "marcador"}]}]
                  }
                ]
              },
              {"type": "pageBreak"},
              {
                "type": "table",
                "content": [
                  {
                    "type": "tableRow",
                    "content": [
                      {
                        "type": "tableCell",
                        "attrs": {"colspan": 2, "rowspan": 1, "colwidth": [96, 96]},
                        "content": [{"type": "paragraph", "content": [{"type": "text", "text": "A"}]}]
                      },
                      {
                        "type": "tableCell",
                        "attrs": {"colspan": 1, "rowspan": 2, "colwidth": [96]},
                        "content": [{"type": "paragraph", "content": [{"type": "text", "text": "B"}]}]
                      }
                    ]
                  },
                  {
                    "type": "tableRow",
                    "content": [
                      {
                        "type": "tableCell",
                        "attrs": {"colspan": 1, "rowspan": 1},
                        "content": [
                          {
                            "type": "paragraph",
                            "content": [
                              {"type": "text", "text": "C"},
                              {
                                "type": "noteRef",
                                "attrs": {"kind": "footnote"},
                                "content": [
                                  {
                                    "type": "paragraph",
                                    "content": [{"type": "text", "text": "Nota na célula"}]
                                  }
                                ]
                              }
                            ]
                          }
                        ]
                      },
                      {
                        "type": "tableCell",
                        "attrs": {
                          "colspan": 1,
                          "rowspan": 1,
                          "borders": "top:double,1.5,#ff0000",
                          "shading": "#eeeeee"
                        },
                        "content": [{"type": "paragraph", "content": [{"type": "text", "text": "D"}]}]
                      }
                    ]
                  }
                ]
              },
              {
                "type": "paragraph",
                "attrs": {"sectionBreak": "s1"},
                "content": [
                  {"type": "text", "text": "Fim da primeira seção"},
                  {
                    "type": "noteRef",
                    "attrs": {"kind": "endnote"},
                    "content": [{"type": "paragraph", "content": [{"type": "text", "text": "Nota de fim"}]}]
                  }
                ]
              },
              {"type": "paragraph", "content": [{"type": "text", "text": "Segunda seção"}]}
            ]
          },
          "styles": {
            "defaults": {
              "paragraph": {"spaceAfter": 8, "lineSpacing": {"kind": "multiple", "factor": 1.0792}},
              "character": {"fontFamily": "Calibri", "fontSize": "11pt"},
              "paragraphStyleId": "Normal",
              "characterStyleId": "DefaultParagraphFont"
            },
            "styles": {
              "Normal": {
                "id": "Normal",
                "name": "Normal",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false
              },
              "DefaultParagraphFont": {
                "id": "DefaultParagraphFont",
                "name": "Default Paragraph Font",
                "type": "character",
                "qFormat": false,
                "hidden": true,
                "custom": false,
                "uiPriority": 1
              },
              "Heading1": {
                "id": "Heading1",
                "name": "heading 1",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "next": "Normal",
                "uiPriority": 9,
                "paragraph": {"spaceBefore": 12, "spaceAfter": 0, "keepNext": true, "keepLines": true, "outlineLevel": 0},
                "character": {"fontSize": "16pt", "color": "#2f5496"}
              },
              "Heading2": {
                "id": "Heading2",
                "name": "heading 2",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "next": "Normal",
                "uiPriority": 9,
                "paragraph": {"spaceBefore": 2, "spaceAfter": 0, "keepNext": true, "keepLines": true, "outlineLevel": 1},
                "character": {"fontSize": "13pt", "color": "#2f5496"}
              },
              "Heading3": {
                "id": "Heading3",
                "name": "heading 3",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "next": "Normal",
                "uiPriority": 9,
                "paragraph": {"spaceBefore": 2, "spaceAfter": 0, "keepNext": true, "keepLines": true, "outlineLevel": 2},
                "character": {"fontSize": "12pt", "color": "#1f3763"}
              },
              "Heading4": {
                "id": "Heading4",
                "name": "heading 4",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "next": "Normal",
                "uiPriority": 9,
                "paragraph": {"spaceBefore": 2, "spaceAfter": 0, "keepNext": true, "keepLines": true, "outlineLevel": 3},
                "character": {"color": "#2f5496", "italic": true}
              },
              "Heading5": {
                "id": "Heading5",
                "name": "heading 5",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "next": "Normal",
                "uiPriority": 9,
                "paragraph": {"spaceBefore": 2, "spaceAfter": 0, "keepNext": true, "keepLines": true, "outlineLevel": 4},
                "character": {"color": "#2f5496"}
              },
              "Heading6": {
                "id": "Heading6",
                "name": "heading 6",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "next": "Normal",
                "uiPriority": 9,
                "paragraph": {"spaceBefore": 2, "spaceAfter": 0, "keepNext": true, "keepLines": true, "outlineLevel": 5},
                "character": {"color": "#1f3763"}
              },
              "ListParagraph": {
                "id": "ListParagraph",
                "name": "List Paragraph",
                "type": "paragraph",
                "qFormat": true,
                "hidden": false,
                "custom": false,
                "basedOn": "Normal",
                "uiPriority": 34,
                "paragraph": {"indentMm": 12.7, "contextualSpacing": true}
              },
              "Hyperlink": {
                "id": "Hyperlink",
                "name": "Hyperlink",
                "type": "character",
                "qFormat": false,
                "hidden": false,
                "custom": false,
                "basedOn": "DefaultParagraphFont",
                "uiPriority": 99,
                "character": {"color": "#0563c1", "underline": true}
              }
            }
          }
        }
        """;
}
