using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>style:font-face</c>: the name a style points to, the family it draws and its generic
/// substitute, the same stack <see cref="Docx.FontTable"/> gives a DOCX font.
/// </summary>
internal sealed class OdtFonts
{
    private readonly Dictionary<string, string> _stacks = new(StringComparer.Ordinal);

    public OdtFonts(params XDocument?[] documents)
    {
        foreach (var face in documents.SelectMany(document =>
                     document?.Root?.Element(Ns.Office + "font-face-decls")?.Elements(Ns.Style + "font-face") ?? []))
        {
            if (face.Attr(Ns.Style + "name") is not { Length: > 0 } name) continue;
            var family = Unquote(face.Attr(Ns.Svg + "font-family") ?? name);
            var generic = GenericOf(face.Attr(Ns.Style + "font-family-generic"));
            _stacks[name] = generic is null ? family : $"{family}, {generic}";
        }
    }

    /// <summary>An undeclared name is the family itself.</summary>
    public string Stack(string name) => _stacks.TryGetValue(name, out var stack) ? stack : name;

    /// <summary>
    /// The declared face that draws <paramref name="stack"/>, the same stack or at least the same
    /// first family, or <c>null</c> when the document declares none.
    /// </summary>
    public string? NameOf(string stack)
    {
        var family = StackOfFamily(stack);
        foreach (var (name, declared) in _stacks)
        {
            if (declared == stack) return name;
        }

        return _stacks.FirstOrDefault(entry => StackOfFamily(entry.Value) == family).Key;
    }

    /// <summary><c>fo:font-family</c>, without a declaration: the first family.</summary>
    public string? StackOfFamily(string families) =>
        families.Split(',').Select(Unquote).FirstOrDefault(family => family.Length > 0);

    private static string Unquote(string value) => value.Trim().Trim('\'', '"').Trim();

    private static string? GenericOf(string? generic) => generic switch
    {
        "swiss" => "sans-serif",
        "roman" => "serif",
        "modern" => "monospace",
        "script" => "cursive",
        "decorative" => "fantasy",
        _ => null,
    };
}
