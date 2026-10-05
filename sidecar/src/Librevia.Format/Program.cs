using Librevia.Format;

// stdout carries the frames: stray text in it would corrupt them, so it goes to stderr.
var frameStream = Console.OpenStandardOutput();
Console.SetOut(Console.Error);

using var lifetime = new CancellationTokenSource();

// Without stopping the loop, a half-written response would arrive as a corrupt frame.
Console.CancelKeyPress += (_, eventArgs) =>
{
    eventArgs.Cancel = true;
    lifetime.Cancel();
};
AppDomain.CurrentDomain.ProcessExit += (_, _) => lifetime.Cancel();

var server = new Server(Console.OpenStandardInput(), frameStream);
await server.RunAsync(lifetime.Token).ConfigureAwait(false);
