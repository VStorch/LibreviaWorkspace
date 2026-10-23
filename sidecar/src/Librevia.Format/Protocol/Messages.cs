using System.Text.Json;
using System.Text.Json.Serialization;

namespace Librevia.Format.Protocol;

public sealed record Request(int Id, string Method, JsonElement Params);

/// <summary>
/// An error for the user, with the rule of <c>src/shared/errors.ts</c>: no stack trace or absolute
/// path. <see cref="Key"/> names the sentence in <c>src/shared/i18n/catalog/sidecar.ts</c>, which
/// main shows in the interface language; <see cref="Detail"/> only goes to the main log. The
/// exceptions the server turns into this payload carry the key as their message.
/// </summary>
public sealed record ErrorPayload(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("key")] string Key,
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
