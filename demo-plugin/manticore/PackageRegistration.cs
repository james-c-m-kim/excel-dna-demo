using System.IO;
using System.Linq;
using Microsoft.Win32;
using NetOffice;

namespace manticore
{
    public class PackageRegistration
    {
        private string registerName;

        public string InstallationPath { get; private set; }

        public PackageRegistration(string productName)
        {
            registerName = productName;
        }

        public bool IsRegistered()
        {
            var app = new NetOffice.ExcelApi.Application();

            // grab an immutable copy of the installed Excel version code
            var versionToken = app.Version;

            var key = Registry.CurrentUser.OpenSubKey($"Software\\Microsoft\\Office\\{versionToken}\\Excel\\Options", true);
            if (key == null) return false;
            var allChildValues = key.GetValueNames();

            // check if the plug-in is already installed...
            var openValues = allChildValues.Where(v => v.StartsWith("OPEN")).ToList();
            var result = openValues.Any(v => key.GetValue(v).ToString().Contains(registerName));

            key.Close();
            return result;
        }

        public void Register()
        {
            var app = new NetOffice.ExcelApi.Application();
            var version = app.Version;
            app.Quit();

            var key = Registry.CurrentUser.OpenSubKey($"Software\\Microsoft\\Office\\{version}\\Excel\\Options", true);
            if (key == null) return;
            var allChildValues = key.GetValueNames();
            // check if the plug-in is already installed...
            var openValues = allChildValues.Where(v => v.StartsWith("OPEN")).ToList();
            if (openValues.Any(v => key.GetValue(v).ToString().Contains(registerName))) return;

            var path = InstallationPath;
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            var xllToRegister = $"{path}\\{registerName}";

            var incr = openValues.Count > 1 ? (openValues.Count - 1).ToString() : "";
            var newOpenValue = $"OPEN{incr}";

            key.SetValue(newOpenValue, xllToRegister);
            key.Close();

        }
    }
}