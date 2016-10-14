using System;

namespace AutomatingMicrosoftExcelUsingDotNet
{
    class Program
    {
        static int Main(string[] args)
        {
            return args.Length < 1 
                ? new Program().ProcessArgs(args) 
                : new Program().ProcessArgs(Console.ReadLine());
        }

        public int ProcessArgs(string[] args)
        {
            return 1;
        }

        public int ProcessArgs(string args)
        {
            return 1;
        }
    }
}
