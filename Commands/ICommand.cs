using AutomatingMicrosoftExcelUsingDotNet.Commands;

namespace AutomatingMicrosoftExcelUsingDotNet.Contracts
{
    public interface ICommand
    {
        int Run(string[] args);
        int Run(BaseOptions options);
    }
}
