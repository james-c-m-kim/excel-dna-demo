using demo_plugin.forms;
using ExcelDna.Integration.CustomUI;

namespace demo_plugin
{
    public static class CtpManager
    {
        private static DemoPane userPaneControl;
        private static DemoPaneHostControl hostControl;
        private static CustomTaskPane excelPane;

        static CtpManager()
        {
            userPaneControl = new DemoPane();
            hostControl = new DemoPaneHostControl();

            excelPane = CustomTaskPaneFactory.CreateCustomTaskPane(hostControl, "User Login");
            excelPane.Visible = false;
            excelPane.Width = 300;
            excelPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        public static void ShowPane()
        {
            excelPane.Visible = true;
        }

        public static void HidePane()
        {
            excelPane.Visible = false;
        }
    }
}