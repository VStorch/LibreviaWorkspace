using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>O espelho de <c>DocumentNode</c> em <c>src/services/document/model.ts</c>: mudam juntos.</summary>
public sealed class Node
{
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    [JsonPropertyName("attrs")]
    public Dictionary<string, JsonNode?>? Attrs { get; set; }

    [JsonPropertyName("content")]
    public List<Node>? Content { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("marks")]
    public List<Mark>? Marks { get; set; }

    public static Node Of(string type, params Node[] children) =>
        new() { Type = type, Content = children.Length == 0 ? null : [.. children] };

    public Node With(string name, JsonNode? value)
    {
        (Attrs ??= [])[name] = value;
        return this;
    }

    /// <summary>
    /// A forma que decide se o usuário mexeu no bloco. Leitor e editor descrevem o
    /// mesmo bloco de jeitos diferentes, e estas diferenças são **de forma**:
    /// <list type="number">
    /// <item>o <c>oid</c> é identidade, e não conteúdo;</item>
    /// <item>atributo nulo é ausente: o ProseMirror materializa todo atributo do schema;</item>
    /// <item>a ordem das marcas é a do schema de um lado e a do <c>w:rPr</c> do outro;</item>
    /// <item>texto vizinho de marcas iguais: o ProseMirror funde, o leitor emite um por <c>w:r</c>;</item>
    /// <item>a ordem das chaves do <c>JsonObject</c>.</item>
    /// </list>
    /// Só para a comparação: quem grava é o XML original.
    /// </summary>
    public string Fingerprint()
    {
        var clone = JsonSerializer.SerializeToNode(this, DocxJson.Options)!;
        Normalize(clone);
        return Canonical(clone)!.ToJsonString();
    }

    private static JsonNode? Canonical(JsonNode? node) => node switch
    {
        JsonObject o => new JsonObject(
            o.OrderBy(entry => entry.Key, StringComparer.Ordinal)
                .Select(entry => KeyValuePair.Create(entry.Key, Canonical(entry.Value?.DeepClone())))),
        JsonArray a => new JsonArray([.. a.Select(item => Canonical(item?.DeepClone()))]),
        _ => node?.DeepClone(),
    };

    private static void Normalize(JsonNode? node)
    {
        switch (node)
        {
            case JsonObject o:
                // A referência de nota vale pelo que aponta: o corpo NotesWriter compara à parte.
                if (o["type"]?.GetValueKind() == JsonValueKind.String &&
                    o["type"]!.GetValue<string>() == "noteRef")
                {
                    o.Remove("content");
                    if (o["attrs"] is JsonObject noteAttrs)
                    {
                        foreach (var entry in noteAttrs.ToList())
                        {
                            if (entry.Key is not ("kind" or "nid" or "mark")) noteAttrs.Remove(entry.Key);
                        }
                    }
                }

                // A equação vale pelo OMML, de que o resto sai; a nova (sem OMML), pelo MathML e pelo modo.
                if (o["type"]?.GetValueKind() == JsonValueKind.String &&
                    o["type"]!.GetValue<string>() == "math" &&
                    o["attrs"] is JsonObject mathAttrs &&
                    mathAttrs["omml"]?.GetValueKind() == JsonValueKind.String &&
                    mathAttrs["omml"]!.GetValue<string>().Length > 0)
                {
                    foreach (var entry in mathAttrs.ToList())
                    {
                        if (entry.Key != "omml") mathAttrs.Remove(entry.Key);
                    }
                }

                if (o["attrs"] is JsonObject attrs)
                {
                    attrs.Remove("oid");
                    // A marca de seção é identidade: o id muda quando uma quebra nova parte a seção.
                    attrs.Remove("sectionBreak");
                    foreach (var entry in attrs.ToList())
                    {
                        if (entry.Value is null) attrs.Remove(entry.Key);
                    }

                    if (attrs.Count == 0) o.Remove("attrs");
                }

                if (o["marks"] is JsonArray marks) SortMarks(marks);
                if (o["content"] is JsonArray content) NormalizeContent(content);
                break;

            case JsonArray a:
                foreach (var item in a) Normalize(item);
                break;
        }
    }

    private static void SortMarks(JsonArray marks)
    {
        foreach (var mark in marks) Normalize(mark);

        // `JsonArray` não ordena no lugar, e um nó só tem um pai: daí o `DeepClone`.
        var sorted = marks
            .Select(mark => mark?.DeepClone())
            .OrderBy(
                mark => (mark as JsonObject)?["type"]?.GetValue<string>() ?? string.Empty,
                StringComparer.Ordinal)
            .ToList();

        marks.Clear();
        foreach (var mark in sorted) marks.Add(mark);
    }

    private static void NormalizeContent(JsonArray content)
    {
        foreach (var child in content) Normalize(child);

        for (var i = content.Count - 1; i > 0; i--)
        {
            if (content[i] is not JsonObject current || content[i - 1] is not JsonObject previous) continue;
            if (!IsText(current) || !IsText(previous)) continue;
            if (!string.Equals(MarksOf(previous), MarksOf(current), StringComparison.Ordinal)) continue;

            previous["text"] = (previous["text"]?.GetValue<string>() ?? string.Empty)
                               + (current["text"]?.GetValue<string>() ?? string.Empty);
            content.RemoveAt(i);
        }
    }

    private static bool IsText(JsonObject node) =>
        string.Equals(node["type"]?.GetValue<string>(), "text", StringComparison.Ordinal);

    private static string MarksOf(JsonObject node) => node["marks"]?.ToJsonString() ?? "null";
}

public sealed class Mark
{
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    [JsonPropertyName("attrs")]
    public Dictionary<string, JsonNode?>? Attrs { get; set; }

    public static Mark Of(string type) => new() { Type = type };

    public static Mark Of(string type, string attribute, JsonNode? value) =>
        new() { Type = type, Attrs = new Dictionary<string, JsonNode?> { [attribute] = value } };
}

public static class DocxJson
{
    public static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };
}
