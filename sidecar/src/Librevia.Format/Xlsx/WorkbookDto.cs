using System.Text.Json;
using System.Text.Json.Serialization;

namespace Librevia.Format.Xlsx;

/// <summary>
/// Mirrors <c>WorkbookModel</c> in <c>src/services/spreadsheet/model.ts</c>: they change together.
/// </summary>
public sealed class WorkbookDto
{
    [JsonPropertyName("sheets")]
    public List<SheetDto> Sheets { get; init; } = [];

    [JsonPropertyName("activeSheet")]
    public int ActiveSheet { get; set; }
}

public sealed class SheetDto
{
    [JsonPropertyName("name")]
    public required string Name { get; set; }

    /// <summary>A sparse map by A1 reference, as in the app model.</summary>
    [JsonPropertyName("cells")]
    public Dictionary<string, CellDto> Cells { get; init; } = [];

    /// <summary>Widths in pixels, by zero-based column index.</summary>
    [JsonPropertyName("columnWidths")]
    public Dictionary<int, double> ColumnWidths { get; init; } = [];

    [JsonPropertyName("rowHeights")]
    public Dictionary<int, double> RowHeights { get; init; } = [];

    [JsonPropertyName("frozenRows")]
    public int FrozenRows { get; set; }

    [JsonPropertyName("frozenColumns")]
    public int FrozenColumns { get; set; }

    [JsonPropertyName("rowCount")]
    public int RowCount { get; set; }

    [JsonPropertyName("columnCount")]
    public int ColumnCount { get; set; }
}

public sealed class CellDto
{
    /// <summary>Number, text or boolean. A date is a serial number.</summary>
    [JsonPropertyName("value")]
    [JsonConverter(typeof(ScalarConverter))]
    public object? Value { get; set; }

    /// <summary>With the leading <c>=</c>, already in the app's language.</summary>
    [JsonPropertyName("formula")]
    public string? Formula { get; set; }

    [JsonPropertyName("style")]
    public CellStyleDto? Style { get; set; }

    /// <summary>By content: that is what spares the cells the user did not touch.</summary>
    public bool SameAs(CellDto? other) =>
        other is not null
        && Equals(Normalize(Value), Normalize(other.Value))
        && Formula == other.Formula
        && CellStyleDto.Same(Style, other.Style);

    /// <summary>
    /// An integer as <c>double</c> and as <c>long</c> differ in <c>Equals</c>, even after JSON.
    /// </summary>
    private static object? Normalize(object? value) => value switch
    {
        null => null,
        bool flag => flag,
        string text => text,
        _ => Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture),
    };
}

/// <summary>
/// A cell value as <c>double</c>, <c>string</c>, <c>bool</c> or nothing. Without it, <c>object?</c>
/// becomes a <c>JsonElement</c>, which does not convert to a number.
/// </summary>
internal sealed class ScalarConverter : JsonConverter<object?>
{
    public override object? Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.True:
                return true;
            case JsonTokenType.False:
                return false;
            case JsonTokenType.String:
                return reader.GetString();
            case JsonTokenType.Number:
                return reader.GetDouble();
            case JsonTokenType.StartObject:
            case JsonTokenType.StartArray:
                // Without skipping, the rest of the JSON would come out misaligned.
                reader.Skip();
                return null;
            default:
                return null;
        }
    }

    public override void Write(Utf8JsonWriter writer, object? value, JsonSerializerOptions options)
    {
        switch (value)
        {
            case null:
                writer.WriteNullValue();
                break;
            case bool flag:
                writer.WriteBooleanValue(flag);
                break;
            case string text:
                writer.WriteStringValue(text);
                break;
            default:
                writer.WriteNumberValue(Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
                break;
        }
    }
}

public sealed class CellStyleDto
{
    [JsonPropertyName("bold")]
    public bool? Bold { get; set; }

    [JsonPropertyName("italic")]
    public bool? Italic { get; set; }

    [JsonPropertyName("underline")]
    public bool? Underline { get; set; }

    [JsonPropertyName("color")]
    public string? Color { get; set; }

    [JsonPropertyName("background")]
    public string? Background { get; set; }

    /// <summary>left, center or right.</summary>
    [JsonPropertyName("align")]
    public string? Align { get; set; }

    /// <summary>general, text, number, currency, percent or date.</summary>
    [JsonPropertyName("format")]
    public string? Format { get; set; }

    [JsonPropertyName("decimals")]
    public int? Decimals { get; set; }

    /// <summary>top, right, bottom, left.</summary>
    [JsonPropertyName("borders")]
    public List<string>? Borders { get; set; }

    public bool IsEmpty =>
        Bold is null && Italic is null && Underline is null && Color is null && Background is null
        && Align is null && Format is null && Decimals is null && (Borders is null || Borders.Count == 0);

    public static bool Same(CellStyleDto? left, CellStyleDto? right)
    {
        if (left is null || left.IsEmpty) return right is null || right.IsEmpty;
        if (right is null) return false;

        return left.Bold == right.Bold
               && left.Italic == right.Italic
               && left.Underline == right.Underline
               && left.Color == right.Color
               && left.Background == right.Background
               && left.Align == right.Align
               && left.Format == right.Format
               && left.Decimals == right.Decimals
               && SameBorders(left.Borders, right.Borders);
    }

    private static bool SameBorders(List<string>? left, List<string>? right)
    {
        var a = left ?? [];
        var b = right ?? [];
        return a.Count == b.Count && !a.Except(b, StringComparer.Ordinal).Any();
    }
}

    /// <summary>The model and what we cannot handle.</summary>
public sealed record XlsxOpenResult(
    [property: JsonPropertyName("workbook")] WorkbookDto Workbook,
    [property: JsonPropertyName("inventory")] Docx.Inventory Inventory);

/// <summary>A spreadsheet error with a sentence ready for the user.</summary>
public sealed class XlsxException(string message) : Exception(message);
