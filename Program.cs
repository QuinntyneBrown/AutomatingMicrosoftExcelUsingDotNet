using AutomatingMicrosoftExcelUsingDotNet.Commands;
using System;

namespace AutomatingMicrosoftExcelUsingDotNet
{
    class Program
    {
        static int Main(string[] args)
        {
            return args.Length < 1 
                ? new Program().ProcessArgs(Console.ReadLine())
                : new Program().ProcessArgs(args);
        }

        public int ProcessArgs(string[] args)
        {
            new WorkbookOpenCommand().Run(args);
            return 1;
        }

        public int ProcessArgs(string args)
        {
            return 1;
        }
    }
}
