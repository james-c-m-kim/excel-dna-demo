using System.Diagnostics;
using System.IO;

namespace manticore
{
    public class PackageRunner
    {
        public void Run()
        {
            var pkgName = "demo-plugin-AddIn.xll";
            var destinationPath = new PackageNameResolver { Organization = "Citi", Product = "DemoPlugin" }.GetPackagePath();
            var fileName = Path.Combine(destinationPath, pkgName);

            var register = new PackageRegistration(pkgName);
            if (register.IsRegistered())
            {
                Process.Start("excel.exe");
            }
            else
            {
                // never been registered before - so start the XLL manually...
                Process.Start(fileName);
            }
        }
    }
}
