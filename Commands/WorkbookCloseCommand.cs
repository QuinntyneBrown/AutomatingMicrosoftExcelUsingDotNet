using System;
using AutomatingMicrosoftExcelUsingDotNet.Contracts;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookCloseCommand : ICommand {
        int Run(ref _Workbook workbook, ref Application xlApp);
    }

    public class WorkbookCloseCommand : IWorkbookCloseCommand
    {
        public int Run(BaseOptions optons)
        {
            throw new NotImplementedException();
        }

        public int Run(string[] args)
        {
            throw new NotImplementedException();
        }

        public int Run(ref _Workbook workbook, ref Application xlApp)
        {
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            xlApp.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(xlApp);
            workbook = null;
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return 1;
        }
            
    }
}
