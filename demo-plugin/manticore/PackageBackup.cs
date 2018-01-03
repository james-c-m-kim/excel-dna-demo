using System;
using System.IO;

namespace manticore
{
    public class PackageBackup
    {
        public void LeaveOneBackup(string destinationPath)
        {
            if (Directory.Exists(destinationPath))
            {
                var di = new DirectoryInfo(destinationPath);
                var parent = di.Parent;

                if (parent != null)
                {
                    var children = parent.GetDirectories();
                    foreach (var child in children)
                    {
                        if (child.Name == di.Name) continue;

                        // if this is a prior backup, delete it...
                        child.Delete(true);
                    }
                }

                // make current folder the new backup
                var backedupVersion = DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss");
                Directory.Move(destinationPath, $"{destinationPath}_{backedupVersion}");
            }

            Directory.CreateDirectory(destinationPath);
        }
    }
}
