using Librevia.Format;

// stdout carrega quadros binários. Um Console.WriteLine perdido no meio do
// código corromperia o fluxo de um jeito difícil de diagnosticar, então tudo
// que for texto vai para stderr — inclusive escrita acidental.
var frameStream = Console.OpenStandardOutput();
Console.SetOut(Console.Error);

using var lifetime = new CancellationTokenSource();

// Encerramento pedido pelo sistema: para o laço em vez de morrer no meio de uma
// resposta pela metade, que o outro lado leria como quadro corrompido.
Console.CancelKeyPress += (_, eventArgs) =>
{
    eventArgs.Cancel = true;
    lifetime.Cancel();
};
AppDomain.CurrentDomain.ProcessExit += (_, _) => lifetime.Cancel();

var server = new Server(Console.OpenStandardInput(), frameStream);
await server.RunAsync(lifetime.Token).ConfigureAwait(false);
