using System;
using AutomatingMicrosoftExcelUsingDotNet.Contracts;
using Microsoft.Office.Interop.Excel;
using static CommandLine.Parser;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookOpenCommand : ICommand {
        _Workbook Run(string workbookPath, ref Application xlApp);
    }


    public class WorkbookOpenCommand : IWorkbookOpenCommand
    {
        public int Run(BaseOptions options)
        {
            new Application().Workbooks.Open(options.WorkbookPath,
                                        0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                                        true, false, 0, true, false, false);
            return 1;
        }

        public _Workbook Run(string workbookPath, ref Application xlApp)
        {            
            _Workbook workbook = xlApp.Workbooks.Open(workbookPath,
                                        0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                                        true, false, 0, true, false, false);
            return workbook;
        }

        public int Run(string[] args)
        {
            var options = new BaseOptions();
            Default.ParseArguments(args, options);
            return Run(options);
        }
    }
}
