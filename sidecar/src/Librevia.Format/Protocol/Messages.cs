using System.Text.Json;
using System.Text.Json.Serialization;

namespace Librevia.Format.Protocol;

public sealed record Request(int Id, string Method, JsonElement Params);

/// <summary>
/// Erro destinado ao usuário final.
/// </summary>
/// <remarks>
/// Mesma regra do lado TypeScript (<c>src/shared/errors.ts</c>): nem stack
/// trace nem caminho absoluto atravessam a fronteira. <see cref="Message"/> é
/// uma frase em português pronta para a tela; <see cref="Detail"/> é técnico e
/// vai só para o log do main.
/// </remarks>
public sealed record ErrorPayload(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("detail")] string? Detail = null);

public sealed record HealthResult(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("version")] string Version,
    [property: JsonPropertyName("runtime")] string Runtime);

public static class JsonOptions
{
    public static readonly JsonSerializerOptions Default = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };
}
