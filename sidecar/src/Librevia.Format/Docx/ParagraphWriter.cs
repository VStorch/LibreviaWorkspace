using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Nó do editor → OOXML. Só roda para blocos que o usuário **editou**.
/// </summary>
/// <remarks>
/// O que este arquivo não souber gerar é perda de verdade, e por isso ele é o
/// lugar mais perigoso da Fase 4. A edição cirúrgica reduz o estrago: blocos
/// intactos nunca passam por aqui — vão direto do arquivo original para o novo.
/// </remarks>
public sealed class ParagraphWriter(MainDocumentPart part, Inventory inventory)
{
    private const int TwipsPerIndentLevel = 720;

    public IEnumerable<OpenXmlElement> Write(Node node, ListContext? list = null)
    {
        switch (node.Type)
        {
            case "paragraph":
            case "heading":
                yield return WriteParagraph(node, list);
                break;

            case "pageBreak":
                yield return new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                break;

            case "table":
                yield return WriteTable(node);
                break;

            case "horizontalRule":
                yield return new Paragraph(new ParagraphProperties(
                    new ParagraphBorders(new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Color = "auto",
                    })));
                break;

            default:
                // Um bloco que não sabemos gerar não pode virar nada em
                // silêncio: vira parágrafo vazio e entra no inventário.
                inventory.NoteLoss($"bloco do tipo \"{node.Type}\"");
                yield return new Paragraph();
                break;
        }
    }

    /// <summary>Numeração herdada do documento, para itens de lista.</summary>
    public sealed record ListContext(int NumberingId, int Level);

    private Paragraph WriteParagraph(Node node, ListContext? list)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();

        if (node.Type == "heading" && AttrInt(node, "level") is { } level)
        {
            properties.ParagraphStyleId = new ParagraphStyleId { Val = "Heading" + level };
        }

        if (AttrString(node, "textAlign") is { } align && align != "left")
        {
            properties.Justification = new Justification
            {
                Val = align switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    _ => JustificationValues.Left,
                },
            };
        }

        if (AttrInt(node, "indent") is { } indent && indent > 0)
        {
            properties.Indentation = new Indentation
            {
                Left = (indent * TwipsPerIndentLevel).ToString(CultureInfo.InvariantCulture),
            };
        }

        if (list is not null)
        {
            properties.NumberingProperties = new NumberingProperties(
                new NumberingLevelReference { Val = list.Level },
                new NumberingId { Val = list.NumberingId });
        }

        if (properties.HasChildren) paragraph.ParagraphProperties = properties;

        foreach (var child in node.Content ?? [])
        {
            foreach (var element in WriteInline(child)) paragraph.AppendChild(element);
        }

        // A quebra volta para onde estava: no fim do parágrafo, dentro de um
        // `w:r`. O leitor a transformou em propriedade do bloco para não pôr um
        // nó de bloco em posição de linha; aqui ela desfaz o caminho. Sem isto,
        // editar o parágrafo que carrega a quebra a apagaria em silêncio.
        if (AttrBool(node, "breakAfter"))
        {
            paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        }

        return paragraph;
    }

    private IEnumerable<OpenXmlElement> WriteInline(Node node)
    {
        switch (node.Type)
        {
            case "text":
                yield return WriteTextRun(node);
                break;

            case "hardBreak":
                yield return new Run(new Break());
                break;

            case "pageBreak":
                yield return new Run(new Break { Type = BreakValues.Page });
                break;

            case "image":
                if (WriteImage(node) is { } image) yield return image;
                break;

            default:
                inventory.NoteLoss($"conteúdo do tipo \"{node.Type}\"");
                break;
        }
    }

    private OpenXmlElement WriteTextRun(Node node)
    {
        var run = new Run();
        var properties = new RunProperties();
        string? hyperlink = null;

        foreach (var mark in node.Marks ?? [])
        {
            switch (mark.Type)
            {
                case "bold": properties.Bold = new Bold(); break;
                case "italic": properties.Italic = new Italic(); break;
                case "strike": properties.Strike = new Strike(); break;
                case "caps": properties.Caps = new Caps(); break;
                case "smallCaps": properties.SmallCaps = new SmallCaps(); break;
                case "underline":
                    properties.Underline = new Underline { Val = UnderlineValues.Single };
                    break;

                case "highlight":
                    if (MarkString(mark, "color") is { } fill)
                    {
                        properties.Shading = new Shading
                        {
                            Val = ShadingPatternValues.Clear,
                            Fill = fill.TrimStart('#'),
                        };
                    }

                    break;

                case "link":
                    hyperlink = MarkString(mark, "href");
                    break;

                case "textStyle":
                    ApplyTextStyle(properties, mark);
                    break;

                default:
                    inventory.NoteLoss($"formatação \"{mark.Type}\"");
                    break;
            }
        }

        if (properties.HasChildren) run.RunProperties = properties;

        // `xml:space="preserve"` senão o Word engole espaço no começo e no fim,
        // e frases coladas aparecem sem separação.
        run.AppendChild(new Text(node.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });

        if (hyperlink is null) return run;

        // `w:hyperlink` embrulha o run — não cabe dentro dele.
        Uri target;
        try
        {
            target = new Uri(hyperlink, UriKind.Absolute);
        }
        catch (UriFormatException)
        {
            inventory.NoteLoss("endereço de link inválido");
            return run;
        }

        var relationship = part.AddHyperlinkRelationship(target, true);
        var link = new Hyperlink { Id = relationship.Id };
        link.AppendChild(run);
        return link;
    }

    private void ApplyTextStyle(RunProperties properties, Mark mark)
    {
        if (MarkString(mark, "color") is { } color)
        {
            properties.Color = new Color { Val = color.TrimStart('#') };
        }

        if (MarkString(mark, "fontFamily") is { } font)
        {
            properties.RunFonts = new RunFonts { Ascii = font, HighAnsi = font };
        }

        if (MarkString(mark, "fontSize") is { } size)
        {
            var digits = size.TrimEnd('p', 't', ' ');
            if (double.TryParse(digits, NumberStyles.Float, CultureInfo.InvariantCulture, out var points))
            {
                // `w:sz` é em meios-pontos.
                var halfPoints = (int)Math.Round(points * 2);
                properties.FontSize = new FontSize { Val = halfPoints.ToString(CultureInfo.InvariantCulture) };
            }
        }
    }

    // --- imagens ------------------------------------------------------------

    private static uint _drawingId = 1000;

    private Run? WriteImage(Node node)
    {
        var source = AttrString(node, "src");
        if (source is null || !source.StartsWith("data:", StringComparison.Ordinal))
        {
            inventory.NoteLoss("imagem sem conteúdo embutido");
            return null;
        }

        var comma = source.IndexOf(',', StringComparison.Ordinal);
        var header = source[5..comma];
        if (!header.EndsWith(";base64", StringComparison.OrdinalIgnoreCase))
        {
            inventory.NoteLoss("imagem em formato não suportado");
            return null;
        }

        var contentType = header[..^";base64".Length];
        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(source[(comma + 1)..]);
        }
        catch (FormatException)
        {
            inventory.NoteLoss("imagem com conteúdo ilegível");
            return null;
        }

        // No OpenXml 3.x `ImagePartType` é classe estática de `PartTypeInfo`,
        // e não mais um enum.
        PartTypeInfo imageType;
        switch (contentType)
        {
            case "image/png": imageType = ImagePartType.Png; break;
            case "image/jpeg": imageType = ImagePartType.Jpeg; break;
            case "image/gif": imageType = ImagePartType.Gif; break;
            case "image/bmp": imageType = ImagePartType.Bmp; break;
            default:
                inventory.NoteLoss($"imagem em {contentType}");
                return null;
        }

        var imagePart = part.AddImagePart(imageType);
        using (var stream = new MemoryStream(bytes))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = part.GetIdOfPart(imagePart);

        // Pixels CSS → EMU: 914400 por polegada, 96 px por polegada.
        var widthPx = AttrInt(node, "width") ?? 600;
        var heightPx = AttrInt(node, "height") ?? (int)Math.Round(widthPx * 0.75);
        var cx = (long)widthPx * 914400 / 96;
        var cy = (long)heightPx * 914400 / 96;

        var id = _drawingId++;

        return new Run(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
            new WordDrawing.Inline(
                new WordDrawing.Extent { Cx = cx, Cy = cy },
                new WordDrawing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new WordDrawing.DocProperties { Id = id, Name = "Imagem " + id },
                new Drawing.Graphic(
                    new Drawing.GraphicData(
                        new Pictures.Picture(
                            new Pictures.NonVisualPictureProperties(
                                new Pictures.NonVisualDrawingProperties { Id = 0U, Name = "Imagem " + id },
                                new Pictures.NonVisualPictureDrawingProperties()),
                            new Pictures.BlipFill(
                                new Drawing.Blip { Embed = relationshipId },
                                new Drawing.Stretch(new Drawing.FillRectangle())),
                            new Pictures.ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = 0L, Y = 0L },
                                    new Drawing.Extents { Cx = cx, Cy = cy }),
                                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                                {
                                    Preset = Drawing.ShapeTypeValues.Rectangle,
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
            }));
    }

    // --- tabelas ------------------------------------------------------------

    private Table WriteTable(Node node)
    {
        var table = new Table();

        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        foreach (var rowNode in node.Content ?? [])
        {
            var row = new TableRow();

            foreach (var cellNode in rowNode.Content ?? [])
            {
                var cell = new TableCell();

                if (AttrInt(cellNode, "colspan") is { } span && span > 1)
                {
                    cell.TableCellProperties = new TableCellProperties(new GridSpan { Val = span });
                }

                var wrote = false;
                foreach (var child in cellNode.Content ?? [])
                {
                    foreach (var element in Write(child))
                    {
                        cell.AppendChild(element);
                        wrote = true;
                    }
                }

                // Célula sem parágrafo torna o documento inválido para o Word.
                if (!wrote) cell.AppendChild(new Paragraph());
                row.AppendChild(cell);
            }

            table.AppendChild(row);
        }

        return table;
    }

    // --- leitura de atributos ----------------------------------------------

    private static bool AttrBool(Node node, string name) =>
        node.Attrs is not null
        && node.Attrs.TryGetValue(name, out var value)
        && value is not null
        && value.GetValueKind() == System.Text.Json.JsonValueKind.True;

    private static string? AttrString(Node node, string name) =>
        node.Attrs is not null && node.Attrs.TryGetValue(name, out var value) && value is not null
            ? value.GetValueKind() == System.Text.Json.JsonValueKind.String ? value.GetValue<string>() : null
            : null;

    private static int? AttrInt(Node node, string name)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue(name, out var value) || value is null) return null;
        return value.GetValueKind() == System.Text.Json.JsonValueKind.Number ? value.GetValue<int>() : null;
    }

    private static string? MarkString(Mark mark, string name) =>
        mark.Attrs is not null && mark.Attrs.TryGetValue(name, out var value) && value is not null
            ? value.GetValueKind() == System.Text.Json.JsonValueKind.String ? value.GetValue<string>() : null
            : null;
}
