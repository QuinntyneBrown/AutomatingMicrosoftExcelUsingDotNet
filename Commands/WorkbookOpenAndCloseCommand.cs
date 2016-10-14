using AutomatingMicrosoftExcelUsingDotNet.Contracts;
using Microsoft.Office.Interop.Excel;
using System;

namespace AutomatingMicrosoftExcelUsingDotNet.Commands
{
    public interface IWorkbookOpenAndCloseCommand : ICommand { }

    public class WorkbookOpenAndCloseCommand : IWorkbookOpenAndCloseCommand
    {
        private IWorkbookOpenCommand _workbookOpenCommand;
        private IWorkbookCloseCommand _workbookCloseCommand;

        public WorkbookOpenAndCloseCommand() {
            _workbookOpenCommand = new WorkbookOpenCommand();
            _workbookCloseCommand = new WorkbookCloseCommand();
        }

        public WorkbookOpenAndCloseCommand(IWorkbookOpenCommand workbookOpenCommand, IWorkbookCloseCommand workbookCloseCommand)
        {
            _workbookCloseCommand = workbookCloseCommand;
            _workbookOpenCommand = workbookOpenCommand;
        }

        public int Run(BaseOptions options)
        {
            throw new NotImplementedException();
        }

        public int Run(string[] args)
        {
            Application xlApp = new Application();
            xlApp.Visible = true;
            _Workbook workbook = _workbookOpenCommand.Run(@"C:\Users\Quinntyne\Documents\WP Narrative.xlsx", ref xlApp);            
            _workbookCloseCommand.Run(ref workbook, ref xlApp);
            return 1;
        }
    }
}
