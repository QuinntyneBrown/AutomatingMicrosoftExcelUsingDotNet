using AutomatingMicrosoftExcelUsingDotNet.Commands;
using Microsoft.Practices.Unity;

namespace AutomatingMicrosoftExcelUsingDotNet
{
    public class UnityConfiguration
    {
        public static IUnityContainer GetContainer() {
            var container = new UnityContainer();
            container.RegisterType<IWorkbookOpenCommand, WorkbookOpenCommand>();
            container.RegisterType<IWorkbookCloseCommand, WorkbookCloseCommand>();
            return container;
        }
    }
}
