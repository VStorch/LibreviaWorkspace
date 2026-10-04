using System.Text.Json;
using System.Text.Json.Serialization;

namespace Librevia.Format.Protocol;

public sealed record Request(int Id, string Method, JsonElement Params);

/// <summary>
/// Erro para o usuário, com a regra de <c>src/shared/errors.ts</c>: sem stack trace
/// nem caminho absoluto. <see cref="Message"/> é a frase da tela; <see cref="Detail"/>
/// vai só ao log do main.
/// </summary>
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
