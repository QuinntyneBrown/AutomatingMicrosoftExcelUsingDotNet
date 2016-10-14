using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookOpenExecuteAndCloseCommand { }

    public class WorkbookOpenExecuteAndCloseCommand: IWorkbookOpenExecuteAndCloseCommand
    {
        
        public WorkbookOpenExecuteAndCloseCommand()
        {

        }

        public int Run(ICollection<Action<Application, _Workbook>> actions, string workbookPath) {
            return 1;
        }
    }
}
