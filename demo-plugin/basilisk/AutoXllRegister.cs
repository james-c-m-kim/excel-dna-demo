using System;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace basilisk
{
    public class AutoXllRegister
    {
        public string registerName = "demo-plugin-Addin.xll";
        public string userPath = "";

        public string InstallationPath
        {
            get
            {
                var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Citi", "DemoPlugIn");
                return path;
            }
        }

        public void RegisterXll()
        {
            var app = (Application)ExcelDnaUtil.Application;

            // grab an immutable copy of the installed Excel version code
            var versionToken = app.Version;

            var key = Registry.CurrentUser.OpenSubKey($"Software\\Microsoft\\Office\\{versionToken}\\Excel\\Options", true);
            if (key == null) return;
            var allChildValues = key.GetValueNames();

            // check if the plug-in is already installed...
            var openValues = allChildValues.Where(v => v.StartsWith("OPEN")).ToList();
            if (openValues.Any(v => key.GetValue(v).ToString().Contains("demo-plugin-AddIn.xll"))) return;

            var path = InstallationPath;
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            var xllToRegister = $"{path}\\demo-plugin-AddIn.xll";

            var incr = openValues.Count > 1 ? (openValues.Count - 1).ToString() : "";
            var newOpenValue = $"OPEN{incr}";

            key.SetValue(newOpenValue, xllToRegister);
            key.Close();
        }
    }
}
