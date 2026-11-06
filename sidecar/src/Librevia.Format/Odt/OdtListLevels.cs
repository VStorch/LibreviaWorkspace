using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>text:list-style</c> → the levels of the editor's <c>numbering</c> (<c>LevelDef</c> in
/// <c>list-numbering.ts</c>): format, number text, start and indent.
/// </summary>
internal static partial class OdtListLevels
{
    public const string Ordered = "orderedList";
    public const string Bullet = "bulletList";

    private const int MaxLevels = ListLevels.Count;
    private const string DefaultBullet = "•";

    private static readonly Dictionary<string, string> Formats = new(StringComparer.Ordinal)
    {
        ["1"] = "decimal",
        ["i"] = "lowerRoman",
        ["I"] = "upperRoman",
        ["a"] = "lowerLetter",
        ["A"] = "upperLetter",
        [""] = "none",
    };

    /// <summary><c>style:num-format</c> → the OOXML name, or <c>null</c> for one the editor lacks.</summary>
    public static string? FormatOf(string? declared) => Formats.GetValueOrDefault(declared ?? string.Empty);

    /// <summary>The kind of each level and the levels, up to the deepest the style declares.</summary>
    public static (List<string> Kinds, JsonArray Levels) Read(XElement? listStyle, Inventory inventory)
    {
        var declared = new Dictionary<int, XElement>();
        foreach (var level in listStyle?.Elements() ?? [])
        {
            if (OdtLength.Integer(level.Attr(Ns.Text + "level")) is { } number and >= 1 and <= MaxLevels)
            {
                declared.TryAdd(number - 1, level);
            }
        }

        var kinds = new List<string>();
        var levels = new JsonArray();
        var deepest = declared.Count == 0 ? -1 : declared.Keys.Max();
        for (var index = 0; index <= deepest; index++)
        {
            if (declared.TryGetValue(index, out var element))
            {
                var ordered = element.Name.LocalName == "list-level-style-number";
                kinds.Add(ordered ? Ordered : Bullet);
                levels.Add(ordered ? Numbered(element, index, inventory) : Bulleted(element));
            }
            else
            {
                kinds.Add(Bullet);
                levels.Add(ListLevels.Defaults(Bullet)[index]!.DeepClone());
            }
        }

        return (kinds, levels);
    }

    private static JsonObject Numbered(XElement element, int index, Inventory inventory)
    {
        var declared = element.Attr(Ns.Style + "num-format") ?? string.Empty;
        if (!Formats.TryGetValue(declared, out var format))
        {
            inventory.NoteLoss(OdtNotices.ListNumberFormat, declared);
            format = "decimal";
        }

        var (indent, hanging) = IndentOf(element);
        return ListLevels.Level(
            format,
            NumberText(element, index, format),
            OdtLength.Integer(element.Attr(Ns.Text + "start-value")) ?? 1,
            indent,
            hanging);
    }

    private static JsonObject Bulleted(XElement element)
    {
        var bullet = element.Attr(Ns.Text + "bullet-char") is { Length: > 0 } declared ? declared : DefaultBullet;
        var (indent, hanging) = IndentOf(element);
        return ListLevels.Level("bullet", ListLevels.Shown(bullet), 1, indent, hanging);
    }

    /// <summary>
    /// LibreOffice's own form (<c>loext:num-list-format="%1%.%2%."</c>) when present; otherwise the
    /// prefix, the levels shown and the suffix.
    /// </summary>
    private static string NumberText(XElement element, int index, string format)
    {
        if (element.Attr(Ns.Loext + "num-list-format") is { Length: > 0 } own) return OwnPlaceholder().Replace(own, "%$1");

        var prefix = element.Attr(Ns.Style + "num-prefix") ?? string.Empty;
        var suffix = element.Attr(Ns.Style + "num-suffix") ?? string.Empty;
        if (format == "none") return prefix + suffix;

        var shown = Math.Clamp(OdtLength.Integer(element.Attr(Ns.Text + "display-levels")) ?? 1, 1, index + 1);
        var numbers = new StringBuilder();
        for (var level = index + 2 - shown; level <= index + 1; level++)
        {
            if (numbers.Length > 0) numbers.Append('.');
            numbers.Append('%').Append(level.ToString(CultureInfo.InvariantCulture));
        }

        return prefix + numbers + suffix;
    }

    /// <summary>The label-alignment mode, or the older space-before and minimum label width.</summary>
    private static (double? Indent, double? Hanging) IndentOf(XElement element)
    {
        var properties = element.Element(Ns.Style + "list-level-properties");
        var alignment = properties?.Element(Ns.Style + "list-level-label-alignment");
        double? indent;
        double? hanging;
        if (alignment is not null)
        {
            indent = OdtLength.Millimetres(alignment.Attr(Ns.Fo + "margin-left"));
            hanging = -OdtLength.Millimetres(alignment.Attr(Ns.Fo + "text-indent"));
        }
        else
        {
            var before = OdtLength.Millimetres(properties?.Attr(Ns.Text + "space-before"));
            hanging = OdtLength.Millimetres(properties?.Attr(Ns.Text + "min-label-width"));
            indent = before is null && hanging is null ? null : (before ?? 0) + (hanging ?? 0);
        }

        return (Rounded(indent), Rounded(hanging));
    }

    private static double? Rounded(double? mm) =>
        mm is { } value ? Math.Round(Math.Max(0, value), OdtLength.MillimetresDecimals) : null;

    [GeneratedRegex("%([0-9]+)%")]
    private static partial Regex OwnPlaceholder();
}
