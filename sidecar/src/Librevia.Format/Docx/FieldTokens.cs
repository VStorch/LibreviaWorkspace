namespace Librevia.Format.Docx;

/// <summary>
/// <c>{n}</c> e <c>{total}</c> no texto de uma faixa: os campos PAGE e NUMPAGES. A
/// mesma gramática para o que se digita e o que <c>HeaderReader.TextOf</c> devolve.
/// </summary>
internal static class FieldTokens
{
    /// <summary>Um trecho de texto, precedido do campo que o abre (nulo no primeiro).</summary>
    public sealed record Segment(string? Field, string Text);

    public static bool Contains(string? text) =>
        text is not null && (text.Contains("{n}", StringComparison.Ordinal) || text.Contains("{total}", StringComparison.Ordinal));

    public static List<Segment> Split(string text)
    {
        var segments = new List<Segment>();
        string? field = null;
        var rest = text;

        while (true)
        {
            var page = rest.IndexOf("{n}", StringComparison.Ordinal);
            var total = rest.IndexOf("{total}", StringComparison.Ordinal);
            var next = new[] { page, total }.Where(at => at >= 0).DefaultIfEmpty(-1).Min();

            if (next < 0)
            {
                segments.Add(new Segment(field, rest));
                return segments;
            }

            segments.Add(new Segment(field, rest[..next]));
            var isPage = next == page;
            field = isPage ? " PAGE " : " NUMPAGES ";
            rest = rest[(next + (isPage ? "{n}" : "{total}").Length)..];
        }
    }
}
