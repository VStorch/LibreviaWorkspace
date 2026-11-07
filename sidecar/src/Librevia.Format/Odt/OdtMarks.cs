using System.Text.Json.Nodes;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A run's effective formatting → editor marks, only what differs from what the block already
/// draws (<paramref name="inherited"/>), in <see cref="RunReader.MarksOf"/>'s order.
/// </summary>
internal static class OdtMarks
{
    public static List<Mark>? Of(CharacterFormat effective, CharacterFormat inherited, string? href, string? charStyle)
    {
        var marks = new List<Mark>();
        if (href is not null) marks.Add(Mark.Of("link", "href", href));

        Toggle(marks, "bold", effective.Bold, inherited.Bold);
        Toggle(marks, "italic", effective.Italic, inherited.Italic);
        Toggle(marks, "strike", effective.Strike, inherited.Strike);
        if (effective.Caps == true && inherited.Caps != true) marks.Add(Mark.Of("caps"));
        if (effective.SmallCaps == true && inherited.SmallCaps != true) marks.Add(Mark.Of("smallCaps"));
        if (effective.Position is CharacterFormat.Superscript or CharacterFormat.Subscript &&
            effective.Position != inherited.Position)
        {
            marks.Add(Mark.Of(effective.Position));
        }

        Toggle(marks, "underline", effective.Underline, inherited.Underline);
        if (effective.Background is { } background && background != CharacterFormat.Transparent &&
            background != inherited.Background)
        {
            marks.Add(Mark.Of("highlight", "color", background));
        }

        if (TextStyleOf(effective, inherited) is { } textStyle) marks.Add(textStyle);
        if (charStyle is not null) marks.Add(Mark.Of("charStyle", "styleId", charStyle));
        return marks.Count == 0 ? null : marks;
    }

    /// <summary>On where the block is off; <see cref="RunReader.Off"/> where the block is on.</summary>
    private static void Toggle(List<Mark> marks, string type, bool? effective, bool? inherited)
    {
        if (effective == true && inherited != true) marks.Add(Mark.Of(type));
        else if (effective == false && inherited == true) marks.Add(RunReader.Off(type));
    }

    private static Mark? TextStyleOf(CharacterFormat effective, CharacterFormat inherited)
    {
        var attributes = new Dictionary<string, JsonNode?>();
        if (effective.Color is { } color && color != inherited.Color) attributes["color"] = color;
        if (effective.FontSizePt is { } size && size != inherited.FontSizePt)
        {
            attributes["fontSize"] = RunReader.FormatPoints(size);
        }

        if (effective.FontFamily is { } font && font != inherited.FontFamily) attributes["fontFamily"] = font;
        return attributes.Count == 0 ? null : new Mark { Type = "textStyle", Attrs = attributes };
    }
}
