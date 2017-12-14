using demo_plugin.forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace demo_plugin
{
    public class DemoAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelAsyncUtil.Initialize();
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => $"EXCEPTION: {ex.ToString()}");

            CtpManager.HidePane();
        }

        public void AutoClose()
        {
        }
    }
}
