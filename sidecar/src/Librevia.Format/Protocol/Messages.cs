using System.Text.Json;
using System.Text.Json.Serialization;

namespace Librevia.Format.Protocol;

public sealed record Request(int Id, string Method, JsonElement Params);

/// <summary>
/// An error for the user, with the rule of <c>src/shared/errors.ts</c>: no stack trace or absolute
/// path. <see cref="Message"/> is the on-screen sentence; <see cref="Detail"/> only goes to the
/// main log.
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
