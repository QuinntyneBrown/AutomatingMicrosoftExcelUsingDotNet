using AutomatingMicrosoftExcelUsingDotNet.Commands;
using Microsoft.Practices.Unity;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutomatingMicrosoftExcelUsingDotNet
{
    public class Program
    {
        private readonly IUnityContainer _container;
        private readonly Dictionary<string, Func<string[], int>> _commands;

        public Program()
        {
            _container = UnityConfiguration.GetContainer();
            _commands = new Dictionary<string, Func<string[], int>>
            {
                ["open"] = _container.Resolve<IWorkbookOpenCommand>().Run,
            };
        }

        static int Main(string[] args)
        {
            return args.Length < 1 
                ? new Program().ProcessArgs(Console.ReadLine())
                : new Program().ProcessArgs(args);
        }

        public int ProcessArgs(string args)
        {
            return 1;
        }

        public int ProcessArgs(string[] args)
        {
            var command = string.Empty;
            var lastArg = 0;
            for (; lastArg < args.Length; lastArg++)
            {
                if (args[lastArg].StartsWith("-"))
                {
                    Console.WriteLine($"Unknown option: {args[lastArg]}");
                }
                else
                {
                    command = args[lastArg];
                    break;
                }
            }

            var appArgs = (lastArg + 1) >= args.Length ? Enumerable.Empty<string>() : args.Skip(lastArg + 1).ToArray();

            int exitCode;
            Func<string[], int> builtIn;
            if (_commands.TryGetValue(command, out builtIn))
            {
                exitCode = builtIn(appArgs.ToArray());
            }
            else
            {
                exitCode = 0;
            }
            return exitCode;
        }

        private static bool IsArg(string candidate, string longName) => IsArg(candidate, shortName: null, longName: longName);

        private static bool IsArg(string candidate, string shortName, string longName)
        {
            return (shortName != null && candidate.Equals("-" + shortName)) || (longName != null && candidate.Equals("--" + longName));
        }
    }
}
