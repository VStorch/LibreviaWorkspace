using System.Globalization;
using DocumentFormat.OpenXml;

namespace Librevia.Format.Docx;

/// <summary>
/// A shape's frame and fill, when CSS can draw them: solid color, solid or dashed stroke. Gradient,
/// texture, image, shadow, 3D and non-rectangular geometry come out with <see cref="Complete"/>
/// false, and only that goes to the warning.
/// </summary>
internal sealed record ShapeLook(
    string? Fill,
    string? Line,
    double LineWidthPt,
    bool Dashed,
    bool Complete)
{
    /// <summary>A shape without declared decoration, and nothing to warn about.</summary>
    internal static readonly ShapeLook Plain = new(null, null, 0, false, true);

    internal bool Draws => Fill is not null || (Line is not null && LineWidthPt > 0);

    private const double EmusPerPoint = Unit.EmusPerPoint;

    /// <summary>
    /// Goes up to whoever has <c>spPr</c>: either the shape or its inner box arrives here. A group
    /// does not mislead, because its own is <c>grpSpPr</c>.
    /// </summary>
    internal static ShapeLook Of(OpenXmlElement element)
    {
        var properties = PropertiesOf(element);
        if (properties is null) return Plain;

        var (fill, fillKnown) = FillOf(properties);
        var (line, width, dashed, lineKnown) = LineOf(properties);

        return new ShapeLook(
            fill,
            line,
            width,
            dashed,
            fillKnown && lineKnown && IsRectangle(properties) && !HasEffects(properties));
    }

    private static OpenXmlElement? PropertiesOf(OpenXmlElement element)
    {
        for (var current = element; current is not null; current = current.Parent)
        {
            var found = current.ChildElements.FirstOrDefault(child => child.LocalName == "spPr");
            if (found is not null) return found;
        }

        return null;
    }

    /// <summary>
    /// Without <c>a:noFill</c> or <c>a:solidFill</c> the shape inherits the theme fill, which we do
    /// not know: the case goes to the warning.
    /// </summary>
    private static (string? Color, bool Known) FillOf(OpenXmlElement properties)
    {
        foreach (var child in properties.ChildElements)
        {
            switch (child.LocalName)
            {
                case "noFill":
                    return (null, true);

                case "solidFill":
                    var color = ColorOf(child);
                    return (color, color is not null);

                case "gradFill":
                case "blipFill":
                case "pattFill":
                case "grpFill":
                    return (null, false);
            }
        }

        return (null, false);
    }

    private static (string? Color, double WidthPt, bool Dashed, bool Known) LineOf(OpenXmlElement properties)
    {
        var line = properties.ChildElements.FirstOrDefault(child => child.LocalName == "ln");

        // Without `a:ln` the shape inherits the style outline, like the fill.
        if (line is null) return (null, 0, false, false);

        var width = Points(Attribute(line, "w"));
        var dashed = line.ChildElements.Any(child => child.LocalName == "prstDash");

        foreach (var child in line.ChildElements)
        {
            switch (child.LocalName)
            {
                case "noFill":
                    return (null, 0, false, true);

                case "solidFill":
                    var color = ColorOf(child);

                    // Zero width with a color is the thinnest stroke, not its absence.
                    return (color, width > 0 ? width : 0.75, dashed, color is not null);

                case "gradFill":
                case "pattFill":
                    return (null, 0, false, false);
            }
        }

        return (null, 0, false, false);
    }

    /// <summary>Without <c>a:prstGeom</c> or <c>a:custGeom</c>, the shape is rectangular.</summary>
    private static bool IsRectangle(OpenXmlElement properties)
    {
        foreach (var child in properties.ChildElements)
        {
            if (child.LocalName == "custGeom") return false;
            if (child.LocalName == "prstGeom") return Attribute(child, "prst") is null or "rect";
        }

        return true;
    }

    private static bool HasEffects(OpenXmlElement properties) =>
        properties.ChildElements.Any(child =>
            child.LocalName is "scene3d" or "sp3d"
            || (child.LocalName == "effectLst" && child.HasChildren)
            || child.LocalName == "effectDag");

    /// <c>a:schemeClr</c> points to the theme: null, and the shape goes to the warning.
    private static string? ColorOf(OpenXmlElement fill)
    {
        var srgb = fill.ChildElements.FirstOrDefault(child => child.LocalName == "srgbClr");
        var value = srgb is null ? null : Attribute(srgb, "val");
        if (value is null || value.Length != 6) return null;

        return "#" + value.ToLowerInvariant();
    }

    private static string? Attribute(OpenXmlElement element, string name) =>
        element.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name).Value;

    private static double Points(string? emus) =>
        double.TryParse(emus, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) && value > 0
            ? Math.Round(value / EmusPerPoint, 2)
            : 0;
}
