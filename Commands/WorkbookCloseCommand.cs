using System;
using AutomatingMicrosoftExcelUsingDotNet.Contracts;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookCloseCommand : ICommand { }

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
    }
}
