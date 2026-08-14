using System.Text;
using System.Text.Json;
using Librevia.Format.Protocol;

namespace Librevia.Format.Tests;

public class FrameIoTests
{
    private static async Task<byte[]> EncodeAsync(object json, byte[] binary)
    {
        using var buffer = new MemoryStream();
        await FrameIo.WriteAsync(
            buffer,
            JsonSerializer.SerializeToUtf8Bytes(json, JsonOptions.Default),
            binary,
            CancellationToken.None);
        return buffer.ToArray();
    }

    [Fact]
    public async Task RoundTripPreservesJsonAndBinary()
    {
        var binary = new byte[] { 0, 255, 13, 10, 0, 127 };
        var encoded = await EncodeAsync(new { id = 1, method = "health" }, binary);

        var frame = await FrameIo.ReadAsync(new MemoryStream(encoded), CancellationToken.None);

        Assert.NotNull(frame);
        Assert.Equal("""{"id":1,"method":"health"}""", Encoding.UTF8.GetString(frame!.Value.Json.Span));
        Assert.Equal(binary, frame.Value.Binary.ToArray());
    }

    [Fact]
    public async Task PreservesBytesThatLookLikeLineEndings()
    {
        // Um protocolo delimitado por \n se despedaçaria aqui. DOCX é ZIP:
        // contém 0x0a, 0x0d e 0x00 o tempo todo.
        var binary = new byte[] { 0x0a, 0x0d, 0x1a, 0x00, 0x50, 0x4b, 0x03, 0x04 };
        var encoded = await EncodeAsync(new { }, binary);

        var frame = await FrameIo.ReadAsync(new MemoryStream(encoded), CancellationToken.None);

        Assert.Equal(binary, frame!.Value.Binary.ToArray());
    }

    [Fact]
    public async Task ReassemblesFrameDeliveredOneByteAtATime()
    {
        // Um pipe entrega o que quiser em cada leitura. Assumir que uma leitura
        // traz a mensagem inteira é o bug clássico desta integração, e ele só
        // aparece com documento grande.
        var binary = new byte[] { 1, 2, 3, 4, 5 };
        var encoded = await EncodeAsync(new { id = 7 }, binary);

        var frame = await FrameIo.ReadAsync(
            new DripStream(encoded, chunkSize: 1),
            CancellationToken.None);

        Assert.Equal("""{"id":7}""", Encoding.UTF8.GetString(frame!.Value.Json.Span));
        Assert.Equal(binary, frame.Value.Binary.ToArray());
    }

    [Fact]
    public async Task ReadsConsecutiveFramesFromOneStream()
    {
        var first = await EncodeAsync(new { id = 1 }, []);
        var second = await EncodeAsync(new { id = 2 }, [9]);

        using var stream = new MemoryStream([.. first, .. second]);

        var a = await FrameIo.ReadAsync(stream, CancellationToken.None);
        var b = await FrameIo.ReadAsync(stream, CancellationToken.None);
        var end = await FrameIo.ReadAsync(stream, CancellationToken.None);

        Assert.Equal("""{"id":1}""", Encoding.UTF8.GetString(a!.Value.Json.Span));
        Assert.Equal("""{"id":2}""", Encoding.UTF8.GetString(b!.Value.Json.Span));
        Assert.Null(end);
    }

    [Fact]
    public async Task SurvivesOneMegabytePayload()
    {
        var binary = new byte[1024 * 1024];
        for (var i = 0; i < binary.Length; i++)
        {
            binary[i] = (byte)(i % 256);
        }

        var encoded = await EncodeAsync(new { id = 1 }, binary);
        var frame = await FrameIo.ReadAsync(
            new DripStream(encoded, chunkSize: 65536),
            CancellationToken.None);

        Assert.Equal(binary, frame!.Value.Binary.ToArray());
    }

    [Fact]
    public async Task ClosedStreamMeansShutdownNotError()
    {
        // stdin fechado é como o main pede para encerrar. Precisa ser uma saída
        // limpa, não uma exceção.
        var frame = await FrameIo.ReadAsync(new MemoryStream([]), CancellationToken.None);

        Assert.Null(frame);
    }

    [Fact]
    public async Task RejectsAbsurdBinaryLengthWithoutAllocating()
    {
        // Sem este teto, um cabeçalho mentiroso nos faria reservar gigabytes.
        var header = new byte[Frame.HeaderBytes];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(
            header.AsSpan(4), Frame.MaxBinaryBytes + 1u);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            FrameIo.ReadAsync(new MemoryStream(header), CancellationToken.None));
    }

    [Fact]
    public async Task RejectsTruncatedFrame()
    {
        var encoded = await EncodeAsync(new { id = 1 }, [1, 2, 3]);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            FrameIo.ReadAsync(
                new MemoryStream(encoded[..^1]),
                CancellationToken.None));
    }

    /// <summary>Entrega poucos bytes por leitura, como um pipe real faz.</summary>
    private sealed class DripStream(byte[] content, int chunkSize) : MemoryStream(content)
    {
        public override int Read(Span<byte> destination) =>
            base.Read(destination[..Math.Min(chunkSize, destination.Length)]);

        public override ValueTask<int> ReadAsync(
            Memory<byte> destination,
            CancellationToken cancellationToken = default) =>
            base.ReadAsync(destination[..Math.Min(chunkSize, destination.Length)], cancellationToken);
    }
}
