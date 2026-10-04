using System.Globalization;
using DocumentFormat.OpenXml;

namespace Librevia.Format.Docx;

/// <summary>
/// A moldura e o preenchimento de uma forma, quando o CSS sabe desenhá-los: cor
/// sólida, traço sólido ou tracejado. Gradiente, textura, imagem, sombra, 3D e
/// geometria que não é retângulo saem com <see cref="Complete"/> falso, e só isso
/// vai ao aviso.
/// </summary>
internal sealed record ShapeLook(
    string? Fill,
    string? Line,
    double LineWidthPt,
    bool Dashed,
    bool Complete)
{
    /// <summary>Forma sem decoração declarada, e sem nada a avisar.</summary>
    internal static readonly ShapeLook Plain = new(null, null, 0, false, true);

    internal bool Draws => Fill is not null || (Line is not null && LineWidthPt > 0);

    private const double EmusPerPoint = Unit.EmusPerPoint;

    /// <summary>
    /// Sobe até quem tem <c>spPr</c>: chega aqui a forma ou a caixa de dentro dela.
    /// O grupo não engana, porque o dele é <c>grpSpPr</c>.
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
    /// Sem <c>a:noFill</c> nem <c>a:solidFill</c> a forma herda o preenchimento do
    /// tema, que não sabemos qual é: o caso vai ao aviso.
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

        // Sem `a:ln` a forma herda o contorno do estilo, como o preenchimento.
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

                    // Espessura zero com cor é o traço mais fino, e não a ausência dele.
                    return (color, width > 0 ? width : 0.75, dashed, color is not null);

                case "gradFill":
                case "pattFill":
                    return (null, 0, false, false);
            }
        }

        return (null, 0, false, false);
    }

    /// <summary>Sem <c>a:prstGeom</c> nem <c>a:custGeom</c>, a forma é retangular.</summary>
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

    /// <summary><c>a:schemeClr</c> aponta o tema: nulo, e a forma vai ao aviso.</summary>
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
