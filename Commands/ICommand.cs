using AutomatingMicrosoftExcelUsingDotNet.Commands;

namespace AutomatingMicrosoftExcelUsingDotNet.Contracts
{
    public interface ICommand
    {
        int run(string[] args);
        int run(BaseOptions optons);
    }
}
