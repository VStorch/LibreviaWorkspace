using System.Globalization;
using DocumentFormat.OpenXml;

namespace Librevia.Format.Docx;

/// <summary>
/// A moldura e o preenchimento de uma forma, quando dá para reproduzi-los.
/// </summary>
/// <remarks>
/// O aviso "moldura e preenchimento de formas" saía em todo documento que
/// tivesse uma caixa de texto, tivesse ela decoração ou não. Nos quatro
/// documentos de evidências do corpus as caixas declaram `a:noFill` e linha de
/// espessura zero — não há moldura nenhuma, e o aviso apontava para uma perda
/// que não existia. Um aviso que aparece sempre é um aviso que se aprende a
/// ignorar, e aí ele deixa de proteger do que importa.
///
/// O que se lê aqui é o caso comum e é o que o CSS sabe desenhar: cor sólida,
/// traço sólido de uma espessura, tracejado. O que sobra —
/// gradiente, textura, imagem de fundo, sombra, três dimensões, e qualquer
/// geometria que não seja um retângulo — sai como <see cref="Complete"/> falso,
/// e é só disso que o inventário passa a falar.
/// </remarks>
internal sealed record ShapeLook(
    string? Fill,
    string? Line,
    double LineWidthPt,
    bool Dashed,
    bool Complete)
{
    /// <summary>Forma sem decoração declarada, e sem nada a avisar.</summary>
    internal static readonly ShapeLook Plain = new(null, null, 0, false, true);

    /// <summary>Há alguma coisa a desenhar?</summary>
    internal bool Draws => Fill is not null || (Line is not null && LineWidthPt > 0);

    /// <summary>1 pt = 12700 EMU.</summary>
    private const double EmusPerPoint = 12700;

    /// <summary>
    /// A decoração da forma que contém este elemento.
    /// </summary>
    /// <remarks>
    /// Sobe pelos ancestrais até achar quem tem `spPr`: o leitor chega aqui
    /// tanto com a forma na mão quanto com a caixa de texto de dentro dela. O
    /// grupo não confunde a busca porque a propriedade dele chama `grpSpPr`.
    /// </remarks>
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
    /// O preenchimento, e se sabemos qual é.
    /// </summary>
    /// <remarks>
    /// Ausência de declaração **não** é ausência de preenchimento: sem `a:noFill`
    /// nem `a:solidFill`, a forma herda o preenchimento do estilo dela, que vem
    /// do tema. Não sabemos qual é, e dizer que não há poria uma caixa branca
    /// onde o documento pede uma azul — então o caso entra no aviso.
    /// </remarks>
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

                    // Espessura zero com cor é o traço mais fino que o formato
                    // sabe pedir, e não a ausência dele: é assim que o filete do
                    // cabeçalho do corpus é gravado.
                    return (color, width > 0 ? width : 0.75, dashed, color is not null);

                case "gradFill":
                case "pattFill":
                    return (null, 0, false, false);
            }
        }

        return (null, 0, false, false);
    }

    /// <summary>
    /// A forma é um retângulo — a única que este desenhista sabe fazer.
    /// </summary>
    /// <remarks>
    /// Sem `a:prstGeom` nem `a:custGeom` a forma é retangular por omissão. Com
    /// `custGeom` ou com um `prstGeom` de outro tipo, desenhar um retângulo
    /// mudaria o desenho, e é disso que o aviso passa a falar.
    /// </remarks>
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

    /// <summary>
    /// A cor de um preenchimento sólido, quando ela está escrita no arquivo.
    /// </summary>
    /// <remarks>
    /// `a:schemeClr` aponta para o tema, e resolver o tema é outro trabalho:
    /// devolve nulo, e a forma entra no aviso em vez de sair com a cor errada.
    /// </remarks>
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
