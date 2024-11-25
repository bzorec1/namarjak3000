using System.Text;

namespace NamarjakPro3000;

public class Program
{
    public static async Task Main(string[] args)
    {
        var indexedArgs = args.Index().ToList();

        var verbose = args.Contains("--verbose");
        var help = args.Contains("-h") || args.Contains("--help");

        string excel = string.Empty;
        string word = string.Empty;

        if (args.Contains("--excel") || args.Contains("-x"))
        {
            var index = args.Index()
                .Where(i => i.Item.Equals("--excel") || i.Item.Equals("-x"))
                .Select(i => i.Index)
                .FirstOrDefault();

            excel = args[index + 1];
        }

        if (args.Contains("--word") || args.Contains("-w"))
        {
            var index = args.Index()
                .Where(i => i.Item.Equals("--word") || i.Item.Equals("-w"))
                .Select(i => i.Index)
                .FirstOrDefault();

            word = args[index + 1];
        }

        while (string.IsNullOrEmpty(excel) || string.IsNullOrEmpty(word))
        {
            Console.Clear();

            if (string.IsNullOrEmpty(excel))
            {
                Console.WriteLine("Please specify a valid excel file.");
                excel = Console.ReadLine()?.Trim('\"') ?? string.Empty;

                if (!string.IsNullOrEmpty(excel) && !File.Exists(excel))
                {
                    Console.WriteLine($"The Excel file '{excel}' does not exist.");
                }
            }

            if (string.IsNullOrEmpty(word))
            {
                Console.WriteLine("Please specify a valid word file.");
                word = Console.ReadLine()?.Trim('\"') ?? string.Empty;

                if (!string.IsNullOrEmpty(word) && !File.Exists(word))
                {
                    Console.WriteLine($"The word file '{word}' does not exist.");
                }
            }
        }

        Console.Clear();

        var cancellationTokenSource = new CancellationTokenSource(TimeSpan.FromHours(1));
        var cancellationToken = cancellationTokenSource.Token;

        StringBuilder stringBuilder = new StringBuilder();

        stringBuilder.AppendLine("NamarjakPro3000");
        stringBuilder.AppendLine("\nPress 'q' to quit.");

        await Console.Out.WriteLineAsync(stringBuilder.ToString());

        Task keyListenerTask = StartKeyListenerAsync(cancellationTokenSource);

        bool running = true;

        while (running)
        {
            stringBuilder.Clear();

            if (cancellationToken.IsCancellationRequested)
            {
                running = false;
                stringBuilder.AppendLine("Cancellation requested. Exiting program...");
            }
            else
            {
                stringBuilder.AppendLine("\nApplication is running... Press 'q' to quit.");
            }

            await Console.Out.WriteLineAsync(stringBuilder.ToString());
            await Task.Delay(1000, cancellationToken);
        }

        await keyListenerTask;
    }

    private static async Task StartKeyListenerAsync(CancellationTokenSource tokenSource)
    {
        await Task.Run(() =>
        {
            while (!tokenSource.Token.IsCancellationRequested)
            {
                if (!Console.KeyAvailable)
                {
                    Thread.Sleep(100);
                    continue;
                }

                ConsoleKeyInfo key = Console.ReadKey(intercept: true);

                if (key.Key != ConsoleKey.Q)
                {
                    Thread.Sleep(100);
                    continue;
                }

                tokenSource.Cancel();
                Console.WriteLine("\n'q' pressed. Requesting cancellation...");

                break;
            }
        });
    }
}