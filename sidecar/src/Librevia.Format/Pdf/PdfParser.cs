using System.Text;

namespace Librevia.Format.Pdf;

/// <summary>
/// Reads objects from a position in the file. Knows the syntax only; finding objects is
/// <see cref="PdfFile"/>'s job, which also resolves a stream length given by reference.
/// </summary>
internal sealed class PdfParser(byte[] data, int position)
{
    private const int MaxDepth = 64;

    public int Position { get; set; } = position;

    public static bool IsWhitespace(byte value) => value is 0 or 9 or 10 or 12 or 13 or 32;

    private static bool IsDelimiter(byte value) => value is (byte)'(' or (byte)')' or (byte)'<' or (byte)'>'
        or (byte)'[' or (byte)']' or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%';

    private static bool IsRegular(byte value) => !IsWhitespace(value) && !IsDelimiter(value);

    public bool AtEnd
    {
        get
        {
            SkipWhitespace();
            return Position >= data.Length;
        }
    }

    public void SkipWhitespace()
    {
        while (Position < data.Length)
        {
            if (IsWhitespace(data[Position]))
            {
                Position++;
            }
            else if (data[Position] == '%')
            {
                while (Position < data.Length && data[Position] is not (byte)'\n' and not (byte)'\r') Position++;
            }
            else
            {
                return;
            }
        }
    }

    /// <summary>A bare word: <c>obj</c>, <c>stream</c>, <c>xref</c>, <c>trailer</c>…</summary>
    public string? PeekKeyword()
    {
        SkipWhitespace();
        var end = Position;
        while (end < data.Length && IsRegular(data[end])) end++;
        return end == Position ? null : Encoding.Latin1.GetString(data, Position, end - Position);
    }

    public void Expect(string keyword)
    {
        if (PeekKeyword() != keyword) throw Broken();
        Position += keyword.Length;
    }

    public long ReadInteger()
    {
        if (ReadObject() is PdfNumber { Integer: { } value }) return value;
        throw Broken();
    }

    public object ReadObject(int depth = 0)
    {
        if (depth > MaxDepth) throw Broken();
        SkipWhitespace();
        if (Position >= data.Length) throw Broken();

        switch (data[Position])
        {
            case (byte)'/':
                return ReadName();
            case (byte)'(':
                return ReadLiteralString();
            case (byte)'[':
                return ReadArray(depth);
            case (byte)'<' when Position + 1 < data.Length && data[Position + 1] == '<':
                return ReadDictionary(depth);
            case (byte)'<':
                return ReadHexString();
            default:
                return ReadNumberOrKeyword();
        }
    }

    /// <summary>After <c>N G obj</c>: the object, and its data when it is a stream.</summary>
    public object ReadIndirectBody(Func<object, long?> resolveLength)
    {
        var value = ReadObject();
        if (value is not PdfDictionary dictionary || PeekKeyword() != "stream") return value;

        Position += "stream".Length;
        if (Position < data.Length && data[Position] == '\r') Position++;
        if (Position < data.Length && data[Position] == '\n') Position++;

        var length = dictionary["Length"] is { } declared ? resolveLength(declared) : null;
        return new PdfStream(dictionary, ReadStreamData(length));
    }

    /// <summary>A wrong <c>/Length</c> is common: then the data ends at <c>endstream</c>.</summary>
    private byte[] ReadStreamData(long? length)
    {
        var start = Position;
        if (length is { } declared && declared >= 0 && start + declared <= data.Length)
        {
            Position = (int)(start + declared);
            if (PeekKeyword() == "endstream") return data[start..(start + (int)declared)];
        }

        var end = IndexOf(data, "endstream"u8, start);
        if (end < 0) throw Broken();
        Position = end;
        var trimmed = end;
        if (trimmed > start && data[trimmed - 1] == '\n') trimmed--;
        if (trimmed > start && data[trimmed - 1] == '\r') trimmed--;
        return data[start..trimmed];
    }

    public static int IndexOf(byte[] haystack, ReadOnlySpan<byte> needle, int from)
    {
        var found = haystack.AsSpan(from).IndexOf(needle);
        return found < 0 ? -1 : from + found;
    }

    private PdfName ReadName()
    {
        Position++;
        var bytes = new List<byte>();
        while (Position < data.Length && IsRegular(data[Position]))
        {
            var value = data[Position++];
            if (value == '#' && Position + 1 < data.Length && TryHex(data[Position], data[Position + 1], out var decoded))
            {
                bytes.Add(decoded);
                Position += 2;
            }
            else
            {
                bytes.Add(value);
            }
        }

        return new PdfName(Encoding.Latin1.GetString([.. bytes]));
    }

    private PdfArray ReadArray(int depth)
    {
        Position++;
        var array = new PdfArray();
        for (;;)
        {
            SkipWhitespace();
            if (Position >= data.Length) throw Broken();
            if (data[Position] == ']')
            {
                Position++;
                return array;
            }

            array.Add(ReadObject(depth + 1));
        }
    }

    private PdfDictionary ReadDictionary(int depth)
    {
        Position += 2;
        var dictionary = new PdfDictionary();
        for (;;)
        {
            SkipWhitespace();
            if (Position + 1 >= data.Length) throw Broken();
            if (data[Position] == '>' && data[Position + 1] == '>')
            {
                Position += 2;
                return dictionary;
            }

            if (ReadObject(depth + 1) is not PdfName key) throw Broken();
            dictionary[key.Value] = ReadObject(depth + 1);
        }
    }

    private PdfString ReadHexString()
    {
        Position++;
        var digits = new List<byte>();
        while (Position < data.Length && data[Position] != '>')
        {
            var value = data[Position++];
            if (!IsWhitespace(value)) digits.Add(value);
        }

        Position++;
        if (digits.Count % 2 == 1) digits.Add((byte)'0');
        var bytes = new byte[digits.Count / 2];
        for (var index = 0; index < bytes.Length; index++)
        {
            if (!TryHex(digits[2 * index], digits[(2 * index) + 1], out bytes[index])) throw Broken();
        }

        return new PdfString(bytes);
    }

    private PdfString ReadLiteralString()
    {
        Position++;
        var bytes = new List<byte>();
        var depth = 1;
        while (Position < data.Length)
        {
            var value = data[Position++];
            switch (value)
            {
                case (byte)'\\':
                    ReadEscape(bytes);
                    continue;
                case (byte)'(':
                    depth++;
                    break;
                case (byte)')' when --depth == 0:
                    return new PdfString([.. bytes]);
                case (byte)'\r':
                    // An unescaped end of line is a line feed, whatever the file used.
                    if (Position < data.Length && data[Position] == '\n') Position++;
                    value = (byte)'\n';
                    break;
            }

            bytes.Add(value);
        }

        throw Broken();
    }

    private void ReadEscape(List<byte> bytes)
    {
        if (Position >= data.Length) return;
        var value = data[Position++];
        switch (value)
        {
            case (byte)'n': bytes.Add((byte)'\n'); return;
            case (byte)'r': bytes.Add((byte)'\r'); return;
            case (byte)'t': bytes.Add((byte)'\t'); return;
            case (byte)'b': bytes.Add((byte)'\b'); return;
            case (byte)'f': bytes.Add((byte)'\f'); return;
            case (byte)'\r':
                if (Position < data.Length && data[Position] == '\n') Position++;
                return;
            case (byte)'\n':
                return;
        }

        if (value is < (byte)'0' or > (byte)'7')
        {
            bytes.Add(value);
            return;
        }

        var octal = value - '0';
        for (var digits = 1; digits < 3 && Position < data.Length && data[Position] is >= (byte)'0' and <= (byte)'7'; digits++)
        {
            octal = (octal * 8) + (data[Position++] - '0');
        }

        bytes.Add((byte)octal);
    }

    private object ReadNumberOrKeyword()
    {
        var start = Position;
        while (Position < data.Length && IsRegular(data[Position])) Position++;
        if (Position == start) throw Broken();

        var word = Encoding.Latin1.GetString(data, start, Position - start);
        switch (word)
        {
            case "true": return true;
            case "false": return false;
            case "null": return PdfNull.Instance;
        }

        if (!IsNumber(word)) throw Broken();
        var number = new PdfNumber(word);
        return number.Integer is { } first && first >= 0 ? ReferenceOr(number, first) : number;
    }

    /// <summary><c>12 0 R</c> reads as three words; only the whole is a reference.</summary>
    private object ReferenceOr(PdfNumber number, long first)
    {
        var mark = Position;
        SkipWhitespace();
        var start = Position;
        while (Position < data.Length && data[Position] is >= (byte)'0' and <= (byte)'9') Position++;
        if (Position > start && Position - start < 10)
        {
            var generation = int.Parse(Encoding.Latin1.GetString(data, start, Position - start));
            SkipWhitespace();
            if (Position < data.Length && data[Position] == 'R'
                && (Position + 1 >= data.Length || !IsRegular(data[Position + 1]))
                && first <= int.MaxValue)
            {
                Position++;
                return new PdfRef((int)first, generation);
            }
        }

        Position = mark;
        return number;
    }

    private static bool IsNumber(string word)
    {
        var digits = 0;
        for (var index = 0; index < word.Length; index++)
        {
            var value = word[index];
            if (char.IsAsciiDigit(value)) digits++;
            else if (value is '+' or '-' && index == 0) continue;
            else if (value != '.') return false;
        }

        return digits > 0;
    }

    private static bool TryHex(byte high, byte low, out byte value)
    {
        var a = HexValue(high);
        var b = HexValue(low);
        value = (byte)((a << 4) | b);
        return a >= 0 && b >= 0;
    }

    private static int HexValue(byte value) => value switch
    {
        >= (byte)'0' and <= (byte)'9' => value - '0',
        >= (byte)'a' and <= (byte)'f' => value - 'a' + 10,
        >= (byte)'A' and <= (byte)'F' => value - 'A' + 10,
        _ => -1,
    };

    public static PdfException Broken() =>
        new("Este PDF está danificado: não foi possível ler a estrutura dele.");
}
