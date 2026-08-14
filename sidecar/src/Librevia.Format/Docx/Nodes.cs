using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// Nó do ProseMirror em forma serializável — o espelho de `DocumentNode` em
/// <c>src/services/document/model.ts</c>. Os dois precisam mudar juntos.
/// </summary>
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
    /// Forma normalizada usada para decidir se o usuário mexeu no bloco.
    /// </summary>
    /// <remarks>
    /// O <c>oid</c> sai fora da comparação: ele é identidade, não conteúdo.
    /// Sem isso, todo bloco pareceria alterado e nada seria preservado — a
    /// edição cirúrgica viraria regeneração completa em silêncio.
    /// </remarks>
    public string Fingerprint()
    {
        var clone = JsonSerializer.SerializeToNode(this, DocxJson.Options)!;
        StripIdentity(clone);
        return clone.ToJsonString();
    }

    private static void StripIdentity(JsonNode? node)
    {
        switch (node)
        {
            case JsonObject o:
                if (o["attrs"] is JsonObject attrs)
                {
                    attrs.Remove("oid");
                    if (attrs.Count == 0)
                    {
                        o.Remove("attrs");
                    }
                }

                foreach (var child in o.ToList())
                {
                    StripIdentity(child.Value);
                }

                break;

            case JsonArray a:
                foreach (var item in a)
                {
                    StripIdentity(item);
                }

                break;
        }
    }
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
