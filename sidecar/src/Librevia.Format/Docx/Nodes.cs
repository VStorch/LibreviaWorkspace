using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// Mirrors <c>DocumentNode</c> in <c>src/services/document/model.ts</c>: they change together.
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
    /// The shape that decides whether the user touched the block. Reader and editor describe the
    /// same block differently, and these differences are **of shape**:
    /// <list type="number">
    /// <item>the <c>oid</c> is identity, not content;</item>
    /// <item>a null attribute is absent: ProseMirror materializes every schema attribute;</item>
    /// <item>mark order is the schema's on one side and the <c>w:rPr</c>'s on the other;</item>
    /// <item>neighbouring text with equal marks: ProseMirror merges, the reader emits one per
    /// <c>w:r</c>;</item>
    /// <item>the key order of the <c>JsonObject</c>.</item>
    /// </list>
    /// Only for comparison: saving writes the original XML.
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
                // A note reference is identified by what it points to: NotesWriter compares the
                // body separately.
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

                // An equation is identified by its OMML, from which the rest derives; a new one
                // (without OMML), by its MathML and mode.
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
                    // The section mark is identity: the id changes when a new break splits the
                    // section.
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

        // `JsonArray` does not sort in place, and a node has a single parent: hence `DeepClone`.
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
