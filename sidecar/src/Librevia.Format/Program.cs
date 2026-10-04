using Librevia.Format;

// O stdout carrega os quadros: texto perdido nele os corromperia, então vai ao stderr.
var frameStream = Console.OpenStandardOutput();
Console.SetOut(Console.Error);

using var lifetime = new CancellationTokenSource();

// Sem parar o laço, a resposta pela metade chegaria como quadro corrompido.
Console.CancelKeyPress += (_, eventArgs) =>
{
    eventArgs.Cancel = true;
    lifetime.Cancel();
};
AppDomain.CurrentDomain.ProcessExit += (_, _) => lifetime.Cancel();

var server = new Server(Console.OpenStandardInput(), frameStream);
await server.RunAsync(lifetime.Token).ConfigureAwait(false);
