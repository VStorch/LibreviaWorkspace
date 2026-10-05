using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <c>w:tblGrid</c>, <c>w:tcBorders</c> and <c>w:shd</c> ↔ node attributes. Each property is
/// **projected** on read and compared on save: equal to the original, the XML is not touched, and
/// an exotic style goes back byte for byte. Only a formatted cell is rewritten, with the
/// approximation in the inventory.
internal static class TableLook
{
    /// <summary>Word's default: 0 top and bottom, 0.075" on the sides.</summary>
    private static readonly int[] WordCellMargins = [0, 108, 0, 108];

    /// <summary>
    /// In twips (top, right, bottom, left), in Word's order: the table, its style and base styles,
    /// the default style and 0/108. With the editor's fixed measure, an eighty-row table gained a
    /// sheet.
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
        // Against a circular `w:basedOn`.
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

    /// <summary>
    /// Only <c>dxa</c> (or untyped) is a measure; <c>nil</c> is zero; percent does not apply to
    /// margins.
    /// </summary>
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

    private const double EighthsPerPoint = 8;

    /// <summary>Without <c>w:sz</c>, half a point, as in Word.</summary>
    private const int DefaultBorderEighths = 4;

    private static readonly string[] Sides = ["top", "right", "bottom", "left"];

    public static int ToPixels(long twips) => (int)Math.Round((double)twips / Unit.TwipsPerPixel);

    public static int ToTwips(int pixels) => pixels * Unit.TwipsPerPixel;


    /// <summary>
    /// From <c>w:tblGrid</c>, which Word uses and the editor mirrors in <c>colgroup</c>; without a
    /// grid, an empty list.
    /// </summary>
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

    /// <summary>Or <c>null</c> without shading.</summary>
    public static string? Shading(TableCellProperties? properties)
    {
        var fill = properties?.Shading?.Fill?.Value;
        if (fill is null || fill.Length != 6) return null;
        if (!int.TryParse(fill, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) return null;

        return "#" + fill.ToLowerInvariant();
    }

    /// <summary>In the canonical text <c>table-format.ts</c> writes.</summary>
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

    /// <summary>
    /// CSS draws five of the twenty-odd; the rest become <c>single</c>, an approximation that only
    /// costs on a formatted cell.
    /// </summary>
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

    /// <c>auto</c> is black, as Word draws it.
    private static string ColorOf(string? value)
    {
        if (value is null || value.Length != 6) return "#000000";
        return int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)
            ? "#" + value.ToLowerInvariant()
            : "#000000";
    }


    /// <summary>If it changed.</summary>
    /// <returns>Whether the element was rewritten.</returns>
    public static bool ApplyShading(TableCellProperties properties, string? model, Inventory inventory)
    {
        if (string.Equals(Shading(properties), model, StringComparison.Ordinal)) return false;

        // A pattern (`pct25`…) has no representation: changing the color flattens it, and the
        // warning says so.
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
    /// Side by side, on the element the cell already has: diagonals, inner borders and
    /// <c>w:space</c> do not exist in the model. An unchanged side is not touched; on a changed
    /// one, only the changed field.
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

        // Only for a side that really changed.
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

    /// <summary>
    /// Through the typed accessors: <c>w:tcBorders</c> is a sequence, and outside it Word refuses.
    /// </summary>
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

    /// <summary>
    /// A theme color goes when the color changes, because in Word it beats <c>w:color</c>.
    /// </summary>
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

    /// <remarks>Zero makes Word ignore the border, and above 255 eighths it leaves the
    /// schema.</remarks>
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
