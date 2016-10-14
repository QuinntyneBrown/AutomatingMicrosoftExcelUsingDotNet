using System;
using AutomatingMicrosoftExcelUsingDotNet.Contracts;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookCloseCommand : ICommand { }

    public class WorkbookCloseCommand : IWorkbookCloseCommand
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
