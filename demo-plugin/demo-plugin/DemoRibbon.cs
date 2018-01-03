using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace demo_plugin
{
    [ComVisible(true)]
    public class DemoRibbon : ExcelRibbon
    {

        public void ShowHello(IRibbonControl rbControl)
        {
            MessageBox.Show("Well there, hello stranger, once again!");
        }

        public void ShowPane(IRibbonControl rbControl)
        {
            CtpManager.ShowPane();
        }
    }
}