using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// <c>w:tblGrid</c>, <c>w:tcBorders</c> e <c>w:shd</c> ↔ atributos do nó. Cada propriedade é
/// **projetada** na leitura e comparada na gravação: igual à do original, o XML não
/// é tocado, e o estilo exótico volta byte a byte. Só a célula formatada é
/// reescrita, com a aproximação no inventário.
/// </summary>
internal static class TableLook
{
    /// <summary>O padrão do Word: 0 em cima e embaixo, 0,075" dos lados.</summary>
    private static readonly int[] WordCellMargins = [0, 108, 0, 108];

    /// <summary>
    /// Em twips (cima, direita, baixo, esquerda), na ordem do Word: a tabela, o
    /// estilo e os estilos-base, o estilo padrão e 0/108. Com a medida fixa do
    /// editor, uma tabela de oitenta linhas ganhava uma folha.
    /// </summary>
    public static int[] CellMargins(Table table, DocumentFormat.OpenXml.Packaging.MainDocumentPart part)
    {
        var sides = new int?[4];
        Fill(sides, table.GetFirstChild<TableProperties>()?.GetFirstChild<TableCellMarginDefault>());

        var styles = part.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(style => style.Type?.Value == StyleValues.Table)
            .ToList() ?? [];
        var styleId = table.GetFirstChild<TableProperties>()?.GetFirstChild<TableStyle>()?.Val?.Value;
        FillFromChain(sides, styles, styleId);
        FillFromChain(sides, styles, styles.FirstOrDefault(style => style.Default?.Value == true)?.StyleId?.Value);

        return [.. sides.Select((side, index) => side ?? WordCellMargins[index])];
    }

    private static void FillFromChain(int?[] sides, List<Style> styles, string? styleId)
    {
        // Contra `w:basedOn` circular.
        for (var hops = 0; styleId is not null && hops < 16; hops++)
        {
            var style = styles.FirstOrDefault(candidate => candidate.StyleId?.Value == styleId);
            if (style is null) return;
            Fill(sides, style.StyleTableProperties?.GetFirstChild<TableCellMarginDefault>());
            styleId = style.BasedOn?.Val?.Value;
        }
    }

    private static void Fill(int?[] sides, TableCellMarginDefault? margins)
    {
        if (margins is null) return;
        sides[0] ??= Twips(margins.TopMargin?.Width?.Value, margins.TopMargin?.Type?.Value);
        sides[1] ??= Twips(margins.TableCellRightMargin?.Width?.Value.ToString(CultureInfo.InvariantCulture), margins.TableCellRightMargin?.Type?.Value)
                     ?? Twips(margins.EndMargin?.Width?.Value, margins.EndMargin?.Type?.Value);
        sides[2] ??= Twips(margins.BottomMargin?.Width?.Value, margins.BottomMargin?.Type?.Value);
        sides[3] ??= Twips(margins.TableCellLeftMargin?.Width?.Value.ToString(CultureInfo.InvariantCulture), margins.TableCellLeftMargin?.Type?.Value)
                     ?? Twips(margins.StartMargin?.Width?.Value, margins.StartMargin?.Type?.Value);
    }

    /// <summary>Só <c>dxa</c> (ou sem tipo) é medida; <c>nil</c> é zero; porcentagem não vale em margem.</summary>
    private static int? Twips(string? width, TableWidthUnitValues? type)
    {
        if (type == TableWidthUnitValues.Nil) return 0;
        if (type is not null && type != TableWidthUnitValues.Dxa) return null;
        return int.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, value)
            : null;
    }

    private static int? Twips(string? width, TableWidthValues? type)
    {
        if (type == TableWidthValues.Nil) return 0;
        if (type is not null && type != TableWidthValues.Dxa) return null;
        return int.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, value)
            : null;
    }

    /// <summary>1440 / 96.</summary>
    private const int TwipsPerPixel = 15;

    private const double EighthsPerPoint = 8;

    /// <summary>Sem <c>w:sz</c>, meio ponto, como no Word.</summary>
    private const int DefaultBorderEighths = 4;

    private static readonly string[] Sides = ["top", "right", "bottom", "left"];

    public static int ToPixels(long twips) => (int)Math.Round((double)twips / TwipsPerPixel);

    public static int ToTwips(int pixels) => pixels * TwipsPerPixel;


    /// <summary>Do <c>w:tblGrid</c>, que o Word usa e o editor espelha no <c>colgroup</c>; sem grade, lista vazia.</summary>
    public static List<int> GridWidths(Table table)
    {
        var widths = new List<int>();

        foreach (var column in table.GetFirstChild<TableGrid>()?.Elements<GridColumn>() ?? [])
        {
            if (!long.TryParse(column.Width?.Value, out var twips) || twips <= 0) return [];
            widths.Add(Math.Max(1, ToPixels(twips)));
        }

        return widths;
    }

    /// <summary>Ou <c>null</c> sem sombreamento.</summary>
    public static string? Shading(TableCellProperties? properties)
    {
        var fill = properties?.Shading?.Fill?.Value;
        if (fill is null || fill.Length != 6) return null;
        if (!int.TryParse(fill, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) return null;

        return "#" + fill.ToLowerInvariant();
    }

    /// <summary>No texto canônico que <c>table-format.ts</c> escreve.</summary>
    public static string? Borders(TableCellProperties? properties)
    {
        var borders = properties?.TableCellBorders;
        if (borders is null) return null;

        var text = new StringBuilder();
        foreach (var side in Sides)
        {
            if (BorderOf(borders, side) is not { } border) continue;
            if (text.Length > 0) text.Append(';');
            text.Append(side).Append(':').Append(Describe(border));
        }

        return text.Length == 0 ? null : text.ToString();
    }

    private static BorderType? BorderOf(TableCellBorders borders, string side) => side switch
    {
        "top" => borders.TopBorder,
        "right" => borders.RightBorder,
        "bottom" => borders.BottomBorder,
        "left" => borders.LeftBorder,
        _ => null,
    };

    private static string Describe(BorderType border)
    {
        var style = StyleOf(border.Val?.Value);
        var eighths = border.Size?.Value ?? DefaultBorderEighths;
        var points = Math.Max(0.25, eighths / EighthsPerPoint);
        var color = ColorOf(border.Color?.Value);

        return $"{style},{points.ToString("0.##", CultureInfo.InvariantCulture)},{color}";
    }

    /// <summary>O CSS desenha cinco dos vinte e tantos; o resto vira <c>single</c>, aproximação que só custa na célula formatada.</summary>
    private static string StyleOf(BorderValues? value)
    {
        if (value is null) return "single";
        if (value == BorderValues.Nil || value == BorderValues.None) return "none";
        if (value == BorderValues.Double) return "double";
        if (value == BorderValues.Dotted) return "dotted";
        if (value == BorderValues.Dashed || value == BorderValues.DashSmallGap) return "dashed";
        return "single";
    }

    private static bool IsExact(BorderValues? value) =>
        value is null
        || value == BorderValues.Nil
        || value == BorderValues.None
        || value == BorderValues.Single
        || value == BorderValues.Double
        || value == BorderValues.Dotted
        || value == BorderValues.Dashed;

    /// <summary><c>auto</c> é preto, como o Word desenha.</summary>
    private static string ColorOf(string? value)
    {
        if (value is null || value.Length != 6) return "#000000";
        return int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)
            ? "#" + value.ToLowerInvariant()
            : "#000000";
    }


    /// <summary>Se ele mudou.</summary>
    /// <returns>Se o elemento foi reescrito.</returns>
    public static bool ApplyShading(TableCellProperties properties, string? model, Inventory inventory)
    {
        if (string.Equals(Shading(properties), model, StringComparison.Ordinal)) return false;

        // A trama (`pct25`…) não tem representação: trocar a cor a achata, e o aviso diz isso.
        var previous = properties.Shading?.Val?.Value;
        if (previous is not null && previous != ShadingPatternValues.Clear && previous != ShadingPatternValues.Nil)
        {
            inventory.NoteLoss("trama de sombreamento de uma célula que você formatou");
        }

        if (model is null)
        {
            properties.Shading = null;
            return true;
        }

        properties.Shading = new Shading
        {
            Val = ShadingPatternValues.Clear,
            Color = "auto",
            Fill = model[1..].ToUpperInvariant(),
        };
        return true;
    }

    /// <summary>
    /// Lado a lado, no elemento que a célula já tem: diagonal, internas e
    /// <c>w:space</c> não existem no modelo. O lado sem mudança não é tocado; no que
    /// mudou, só o campo que mudou.
    /// </summary>
    public static bool ApplyBorders(TableCellProperties properties, string? model, Inventory inventory)
    {
        if (string.Equals(Borders(properties), model, StringComparison.Ordinal)) return false;

        var wanted = ParseSides(model);
        var borders = properties.TableCellBorders ?? new TableCellBorders();
        var approximated = false;

        foreach (var side in Sides)
        {
            var current = BorderOf(borders, side);
            wanted.TryGetValue(side, out var target);

            if (current is not null && target is not null && Describe(current) == string.Join(',', target)) continue;
            if (current is null && target is null) continue;

            if (current is not null && !IsExact(current.Val?.Value)) approximated = true;

            if (target is null)
            {
                SetSide(borders, side, null);
                continue;
            }

            if (current is null)
            {
                SetSide(borders, side, side switch
                {
                    "top" => Build<TopBorder>(target),
                    "right" => Build<RightBorder>(target),
                    "bottom" => Build<BottomBorder>(target),
                    _ => Build<LeftBorder>(target),
                });
                continue;
            }

            Update(current, target);
        }

        // Só para o lado que de fato mudou.
        if (approximated) inventory.NoteLoss("estilo de borda de uma célula que você formatou");

        properties.TableCellBorders = borders.HasChildren ? borders : null;
        return true;
    }

    private static Dictionary<string, string[]> ParseSides(string? model)
    {
        var sides = new Dictionary<string, string[]>(StringComparer.Ordinal);
        foreach (var part in (model ?? "").Split(';'))
        {
            var colon = part.IndexOf(':', StringComparison.Ordinal);
            if (colon <= 0) continue;

            var pieces = part[(colon + 1)..].Split(',');
            if (pieces.Length < 3 || !Sides.Contains(part[..colon])) continue;
            sides[part[..colon]] = pieces;
        }

        return sides;
    }

    /// <summary>Pelos acessores com tipo: o <c>w:tcBorders</c> é uma sequência, e fora dela o Word recusa.</summary>
    private static void SetSide(TableCellBorders borders, string side, BorderType? border)
    {
        switch (side)
        {
            case "top": borders.TopBorder = (TopBorder?)border; break;
            case "right": borders.RightBorder = (RightBorder?)border; break;
            case "bottom": borders.BottomBorder = (BottomBorder?)border; break;
            case "left": borders.LeftBorder = (LeftBorder?)border; break;
        }
    }

    /// <summary>A cor de tema sai quando a cor muda, porque no Word ela vence o <c>w:color</c>.</summary>
    private static void Update(BorderType border, string[] target)
    {
        var (style, width, color) = (target[0], target[1], target[2]);
        var before = Describe(border).Split(',');

        if (before[0] != style) border.Val = ValueOf(style);
        if (before[1] != width) border.Size = SizeOf(width);
        if (before[2] != color)
        {
            if (color.Length == 7 && color[0] == '#') border.Color = color[1..].ToUpperInvariant();
            border.ThemeColor = null;
            border.ThemeTint = null;
            border.ThemeShade = null;
        }
    }

    private static BorderValues ValueOf(string style) => style switch
    {
        "none" => BorderValues.Nil,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        _ => BorderValues.Single,
    };

    /// <remarks>Zero faz o Word ignorar a borda, e acima de 255 oitavos sai do esquema.</remarks>
    private static uint SizeOf(string width)
    {
        var points = double.TryParse(width, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0.5;
        return (uint)Math.Clamp(Math.Round(points * EighthsPerPoint), 2, 255);
    }

    private static TBorder Build<TBorder>(string[] pieces) where TBorder : BorderType, new()
    {
        var (style, width, color) = (pieces[0], pieces[1], pieces[2]);
        var border = new TBorder { Val = ValueOf(style), Size = SizeOf(width) };
        if (color.Length == 7 && color[0] == '#') border.Color = color[1..].ToUpperInvariant();
        return border;
    }
}
