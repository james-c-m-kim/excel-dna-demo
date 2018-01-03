using System;
using System.IO;

namespace manticore
{
    public class PackageNameResolver
    {
        public string Organization { get; set; }
        public string Product { get; set; }

        public string GetPackagePath()
        {
            var destinationPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                Organization, Product);

            return destinationPath;
        }
    }
}
