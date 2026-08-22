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
    /// Compara o que o leitor produziu com o que volta do editor. Os dois
    /// descrevem o mesmo bloco de jeitos diferentes, e as diferenças abaixo são
    /// **de forma, não de conteúdo** — tratá-las como edição faz a gravação
    /// cirúrgica regenerar o documento inteiro em silêncio, que é o risco nº 1
    /// do plano técnico.
    ///
    /// Medido em <c>modelo-de-manual.docx</c>, abrindo e salvando sem
    /// editar nada: dos 15 blocos, batiam 2. Cada normalização foi acrescentada
    /// olhando o que ainda sobrava — 2 → 9 → 11 → 15.
    ///
    /// <list type="number">
    /// <item><b>identidade</b>: o <c>oid</c> é quem o bloco é, não o que ele diz.</item>
    /// <item><b>atributo nulo</b>: o ProseMirror materializa <b>todo</b> atributo
    /// declarado no schema, inclusive os que o documento não menciona, e devolve
    /// <c>null</c> neles; o leitor simplesmente não os escreve. Ausente e nulo
    /// são a mesma afirmação — "o documento não diz nada sobre isto".</item>
    /// <item><b>ordem das marcas</b>: o ProseMirror ordena as marcas pela posição
    /// delas no schema, o leitor as emite na ordem em que leu o <c>w:rPr</c>.
    /// Negrito antes ou depois da cor é a mesma formatação.</item>
    /// <item><b>texto vizinho</b>: o leitor emite um nó por <c>w:r</c>, porque é
    /// assim que o arquivo está escrito — "Acme® Software" chega partido
    /// quando o <c>®</c> tem <c>rPr</c> próprio, ainda que igual ao do vizinho.
    /// O ProseMirror funde nós de texto com marcas iguais; é invariante do
    /// modelo dele, não escolha nossa.</item>
    /// <item><b>ordem das chaves</b>: <c>JsonObject</c> preserva a ordem de
    /// inserção, e os dois lados montam os atributos em ordens diferentes —
    /// <c>{lineHeight, fontSize}</c> de um lado, <c>{fontSize, lineHeight}</c>
    /// do outro. Sozinha, esta diferença já reprovava <b>todos</b> os blocos, e
    /// por isso as outras normalizações não mostravam efeito até esta entrar.</item>
    /// </list>
    ///
    /// Vale só para a comparação. O que vai para a tela continua separado por
    /// run, e quem grava é o XML original — não isto.
    /// </remarks>
    public string Fingerprint()
    {
        var clone = JsonSerializer.SerializeToNode(this, DocxJson.Options)!;
        Normalize(clone);
        return Canonical(clone)!.ToJsonString();
    }

    /// <summary>Mesma árvore, chaves em ordem estável.</summary>
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
                if (o["attrs"] is JsonObject attrs)
                {
                    attrs.Remove("oid");
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

    /// <summary>Marcas em ordem estável, para negrito-antes-de-cor não diferir de cor-antes-de-negrito.</summary>
    private static void SortMarks(JsonArray marks)
    {
        foreach (var mark in marks) Normalize(mark);

        // Cópia e reinserção: `JsonArray` não ordena no lugar, e um nó só pode
        // ter um pai — daí o `DeepClone` antes de limpar.
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
