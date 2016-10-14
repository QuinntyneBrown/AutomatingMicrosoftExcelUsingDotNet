using System;
using AutomatingMicrosoftExcelUsingDotNet.Contracts;
using Microsoft.Office.Interop.Excel;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookOpenCommand : ICommand { }

    public class WorkbookOpenCommand : IWorkbookOpenCommand
    {
        public int run(BaseOptions optons)
        {
            throw new NotImplementedException();
        }

        public int run(string[] args)
        {
            throw new NotImplementedException();
        }
    }
}
