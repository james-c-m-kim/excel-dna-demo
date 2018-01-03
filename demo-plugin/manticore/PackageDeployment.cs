using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;

namespace manticore
{
    public class PackageDeployment
    {
        public void Deploy()
        {
            // get current exe folder
            var location = Assembly.GetAssembly(typeof(PackageDeployment)).Location;
            var uri = new Uri(location);
            var fullName = uri.LocalPath;
            var fi = new FileInfo(fullName);
            var currentPath = fi.DirectoryName;

            var zips = new DirectoryInfo(currentPath).GetFiles("*.zip"); 

            // for now, just treating as if the deployer can handle one zip only.
            var zip = zips.FirstOrDefault();

            if (zip == null) return;

            // deploy target to local app data citi plugin
            var destinationPath = new PackageNameResolver { Organization = "Citi", Product = "DemoPlugin" }.GetPackagePath();
            new PackageBackup().LeaveOneBackup(destinationPath);
            ZipFile.ExtractToDirectory(zip.FullName, destinationPath);
        }
    }
}
