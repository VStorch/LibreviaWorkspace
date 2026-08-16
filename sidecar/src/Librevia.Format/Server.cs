using System.Reflection;
using System.Text.Json;
using Librevia.Format.Protocol;

namespace Librevia.Format;

/// <summary>
/// O laço stdio: lê um quadro, atende, responde. Um pedido por vez.
/// </summary>
/// <remarks>
/// Serial de propósito. Paralelizar aqui traria concorrência para dentro de um
/// processo cujo trabalho é mexer em documento — e o ganho seria nenhum, porque
/// o usuário abre um arquivo de cada vez. Quem cuida de prazo e de cancelamento
/// é o lado TypeScript, que derruba o processo se ele demorar.
/// </remarks>
public sealed class Server(Stream input, Stream output)
{
    private readonly Dictionary<string, Func<Request, ReadOnlyMemory<byte>, CancellationToken, Task<Reply>>>
        _handlers = new(StringComparer.Ordinal)
        {
            ["health"] = static (_, _, _) => Task.FromResult(Reply.Of(Health())),
            // Existe para provar, de ponta a ponta, que binário grande atravessa
            // inteiro. Não toca em disco e não guarda estado.
            ["diagnostics.echo"] = static (_, binary, _) => Task.FromResult(new Reply(null, binary)),
            ["docx.open"] = static (_, binary, _) =>
                Task.FromResult(Reply.Of(Docx.DocxReader.Read(binary.ToArray()))),
            // O binário são os bytes originais que o main guardou na abertura;
            // o modelo vem nos parâmetros. O sidecar não guarda nada entre um
            // pedido e outro — ver docs/02-docx-cirurgico.md.
            ["docx.save"] = static (request, binary, _) =>
            {
                var model = request.Params.Deserialize<Docx.DocumentModelDto>(JsonOptions.Default)
                            ?? throw new Docx.DocxException("O documento a gravar chegou vazio.");
                var (bytes, result) = Docx.DocxWriter.Write(binary.ToArray(), model);
                return Task.FromResult(new Reply(result, bytes));
            },
            ["xlsx.open"] = static (_, binary, _) =>
                Task.FromResult(Reply.Of(Xlsx.XlsxReader.Read(binary.ToArray()))),
            // Mesmo contrato do docx.save: os bytes originais entram pelo
            // binário e o modelo pelos parâmetros. Binário vazio quer dizer
            // planilha nova, sem original para preservar.
            ["xlsx.save"] = static (request, binary, _) =>
            {
                var model = request.Params.Deserialize<Xlsx.WorkbookDto>(JsonOptions.Default)
                            ?? throw new Xlsx.XlsxException("A planilha a gravar chegou vazia.");
                var original = binary.IsEmpty ? null : binary.ToArray();
                var (bytes, result) = Xlsx.XlsxWriter.Write(original, model);
                return Task.FromResult(new Reply(result, bytes));
            },
        };

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
                // Fluxo corrompido: não há como saber onde o próximo quadro
                // começa, então sair é mais honesto que adivinhar. O main
                // percebe a saída e sobe um processo limpo no próximo pedido.
                await Console.Error.WriteLineAsync($"quadro inválido: {problem.Message}").ConfigureAwait(false);
                return;
            }

            // stdin fechado — é assim que o main pede para encerrar.
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
                      ?? throw new InvalidDataException("pedido vazio");

            if (!_handlers.TryGetValue(request.Method, out var handler))
            {
                await RespondErrorAsync(
                    request.Id,
                    new ErrorPayload("UNKNOWN_METHOD", "O aplicativo pediu uma operação que este serviço não conhece.",
                        request.Method),
                    cancellation).ConfigureAwait(false);
                return;
            }

            var reply = await handler(request, frame.Binary, cancellation).ConfigureAwait(false);
            await RespondOkAsync(request.Id, reply, cancellation).ConfigureAwait(false);
        }
        catch (Docx.DocxException problem)
        {
            // Já traz frase pronta para o usuário — não vira "erro inesperado".
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
        catch (Exception problem) when (problem is not OperationCanceledException)
        {
            // Qualquer falha inesperada vira uma resposta, nunca um processo
            // morto em silêncio: o main precisa de algo para mostrar ao usuário.
            await Console.Error.WriteLineAsync(problem.ToString()).ConfigureAwait(false);
            await RespondErrorAsync(
                request?.Id ?? 0,
                new ErrorPayload("INTERNAL", "Ocorreu um erro inesperado ao processar o documento.",
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
        // Reportar as versões das bibliotecas OOXML deixa de ser curiosidade
        // quando um documento abre errado só numa máquina.
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
