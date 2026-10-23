using System.Reflection;
using System.Text.Json;
using Librevia.Format.Protocol;

namespace Librevia.Format;

/// <summary>
/// The stdio loop, one request at a time: people open one file at a time, and timeouts and
/// cancellation belong to the TypeScript side, which kills the process.
/// </summary>
public sealed class Server(Stream input, Stream output)
{
    private readonly Dictionary<string, Func<Request, ReadOnlyMemory<byte>, CancellationToken, Task<Reply>>>
        _handlers = new(StringComparer.Ordinal)
        {
            ["health"] = static (_, _, _) => Task.FromResult(Reply.Of(Health())),
            // Proves a large binary gets through whole.
            ["diagnostics.echo"] = static (_, binary, _) => Task.FromResult(new Reply(null, binary)),
            ["docx.open"] = static (_, binary, _) =>
                Task.FromResult(Reply.Of(Docx.DocxReader.Read(binary.ToArray()))),
            // The binary is the original bytes main kept; the sidecar keeps no state.
            ["docx.save"] = static (request, binary, _) =>
            {
                var model = request.Params.Deserialize<Docx.DocumentModelDto>(JsonOptions.Default)
                            ?? throw new Docx.DocxException("sidecar.error.documentMissing");
                var (bytes, result) = Docx.DocxWriter.Write(binary.ToArray(), model);
                return Task.FromResult(new Reply(result, bytes));
            },
            ["docx.create"] = static (request, _, _) =>
            {
                var create = request.Params.Deserialize<Docx.DocxCreateDto>(JsonOptions.Default)
                             ?? throw new Docx.DocxException("sidecar.error.pageMissing");
                return Task.FromResult(new Reply(null, Docx.DocxTemplate.Create(create.Page, create.Styles)));
            },
            ["xlsx.open"] = static (_, binary, _) =>
                Task.FromResult(Reply.Of(Xlsx.XlsxReader.Read(binary.ToArray()))),
            // An empty binary is a new spreadsheet, without an original.
            ["xlsx.save"] = static (request, binary, _) =>
            {
                var model = request.Params.Deserialize<Xlsx.WorkbookDto>(JsonOptions.Default)
                            ?? throw new Xlsx.XlsxException("sidecar.error.workbookMissing");
                var original = binary.IsEmpty ? null : binary.ToArray();
                var (bytes, result) = Xlsx.XlsxWriter.Write(original, model);
                return Task.FromResult(new Reply(result, bytes));
            },
            ["pdf.inspect"] = static (_, binary, _) =>
                Task.FromResult(Reply.Of(Pdf.PdfMetadata.Inspect(binary.ToArray()))),
            // Only the bytes to append: main writes them after the original, which stays as it was.
            ["pdf.setMetadata"] = static (request, binary, _) =>
            {
                var patch = request.Params.Deserialize<Docx.PropertiesDto>(JsonOptions.Default)
                            ?? throw new Pdf.PdfException("sidecar.error.propertiesMissing");
                var update = Pdf.PdfMetadata.Update(binary.ToArray(), patch);
                return Task.FromResult(new Reply(new { appended = update.Length }, update));
            },
            ["package.inspect"] = static (request, binary, _) =>
                Task.FromResult(Reply.Of(Privacy.PersonalData.Inspect(binary.ToArray(), PrivacyRequestOf(request)))),
            // The whole clean package: unlike a PDF, a zip cannot grow by appending.
            ["package.scrub"] = static (request, binary, _) =>
            {
                var (bytes, changes) = Privacy.PersonalData.Scrub(binary.ToArray(), PrivacyRequestOf(request));
                return Task.FromResult(new Reply(new { changes }, bytes));
            },
        };

    private static Privacy.PrivacyRequest PrivacyRequestOf(Request request) =>
        request.Params.Deserialize<Privacy.PrivacyRequest>(JsonOptions.Default)
        ?? throw new Privacy.PrivacyException("sidecar.error.inspectionMissing");

    public sealed record Reply(object? Result, ReadOnlyMemory<byte> Binary)
    {
        public static Reply Of(object? result) => new(result, ReadOnlyMemory<byte>.Empty);
    }

    public async Task RunAsync(CancellationToken cancellation)
    {
        while (!cancellation.IsCancellationRequested)
        {
            Frame? frame;
            try
            {
                frame = await FrameIo.ReadAsync(input, cancellation).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (InvalidDataException problem)
            {
                // Without knowing where the next frame starts, exit; main starts another process.
                await Console.Error.WriteLineAsync($"invalid frame: {problem.Message}").ConfigureAwait(false);
                return;
            }

            if (frame is null)
            {
                return;
            }

            await HandleAsync(frame.Value, cancellation).ConfigureAwait(false);
        }
    }

    private async Task HandleAsync(Frame frame, CancellationToken cancellation)
    {
        Request? request = null;
        try
        {
            request = JsonSerializer.Deserialize<Request>(frame.Json.Span, JsonOptions.Default)
                      ?? throw new InvalidDataException("empty request");

            if (!_handlers.TryGetValue(request.Method, out var handler))
            {
                await RespondErrorAsync(
                    request.Id,
                    new ErrorPayload("UNKNOWN_METHOD", "sidecar.error.unknownMethod",
                        request.Method),
                    cancellation).ConfigureAwait(false);
                return;
            }

            var reply = await handler(request, frame.Binary, cancellation).ConfigureAwait(false);
            await RespondOkAsync(request.Id, reply, cancellation).ConfigureAwait(false);
        }
        catch (Docx.DocxException problem)
        {
            await RespondErrorAsync(
                request?.Id ?? 0,
                new ErrorPayload("DOCX_INVALID", problem.Message),
                cancellation).ConfigureAwait(false);
        }
        catch (Xlsx.XlsxException problem)
        {
            await RespondErrorAsync(
                request?.Id ?? 0,
                new ErrorPayload("XLSX_INVALID", problem.Message),
                cancellation).ConfigureAwait(false);
        }
        catch (Pdf.PdfException problem)
        {
            await RespondErrorAsync(
                request?.Id ?? 0,
                new ErrorPayload("PDF_INVALID", problem.Message),
                cancellation).ConfigureAwait(false);
        }
        catch (Privacy.PrivacyException problem)
        {
            await RespondErrorAsync(
                request?.Id ?? 0,
                new ErrorPayload("PACKAGE_INVALID", problem.Message),
                cancellation).ConfigureAwait(false);
        }
        catch (Exception problem) when (problem is not OperationCanceledException)
        {
            // Main needs a response to show, not a dead process.
            await Console.Error.WriteLineAsync(problem.ToString()).ConfigureAwait(false);
            await RespondErrorAsync(
                request?.Id ?? 0,
                new ErrorPayload("INTERNAL", "sidecar.error.unexpected",
                    problem.GetType().Name),
                cancellation).ConfigureAwait(false);
        }
    }

    private Task RespondOkAsync(int id, Reply reply, CancellationToken cancellation) =>
        WriteAsync(new { id, ok = true, result = reply.Result }, reply.Binary, cancellation);

    private Task RespondErrorAsync(int id, ErrorPayload error, CancellationToken cancellation) =>
        WriteAsync(new { id, ok = false, error }, ReadOnlyMemory<byte>.Empty, cancellation);

    private Task WriteAsync(object message, ReadOnlyMemory<byte> binary, CancellationToken cancellation) =>
        FrameIo.WriteAsync(
            output,
            JsonSerializer.SerializeToUtf8Bytes(message, JsonOptions.Default),
            binary,
            cancellation);

    private static HealthResult Health() => new(
        Name: "Librevia.Format",
        Version: typeof(Server).Assembly.GetName().Version?.ToString() ?? "0.0.0",
        // For the document that opens wrong on a single machine.
        Runtime: string.Join(' ',
            Environment.Version.ToString(),
            $"OpenXml={VersionOf("DocumentFormat.OpenXml")}",
            $"ClosedXML={VersionOf("ClosedXML")}"));

    private static string VersionOf(string assemblyName)
    {
        try
        {
            return Assembly.Load(assemblyName).GetName().Version?.ToString() ?? "?";
        }
        catch (Exception problem) when (problem is FileNotFoundException or BadImageFormatException)
        {
            return "ausente";
        }
    }
}
