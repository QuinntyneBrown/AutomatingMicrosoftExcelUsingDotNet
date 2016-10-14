using CommandLine;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public class BaseOptions
    {
        [Option("workbookpath", Required = false, HelpText = "Workbook Path")]
        public string WorkbookPath { get; set; }
    }
}
