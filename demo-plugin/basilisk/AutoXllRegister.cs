using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace basilisk
{
    public class AutoXllRegister
    {
        public string registerName = "demo-plugin-Addin.xll";
        public string userPath = "";

        public void RegisterXll()
        {
            var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Options", true);
            var allChildValues = key.GetValueNames();

            var openValues = allChildValues.Where(v => v.StartsWith("OPEN")).ToList();

            if (openValues.Any(v => key.GetValue(v).ToString().Contains("demo-plugin-AddIn.xll"))) return;


            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Citi", "DemoPlugIn");

            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            var xllToRegister = $"{path}\\demo-plugin-AddIn.xll";

            var incr = openValues.Count > 1 ? (openValues.Count - 1).ToString() : "";
            var newOpenValue = $"OPEN{incr}";

            key.SetValue(newOpenValue, xllToRegister);
            key.Close();
        }
    }
}
