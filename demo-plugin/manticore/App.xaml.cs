using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace manticore
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            try
            {
                // unpack/deploy the XLL
                new PackageDeployment().Deploy();

                // execute the XLL for first time
                new PackageRunner().Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to deploy:\n{ex.Message}");
            }
            finally
            {
                Environment.Exit(0);
            }
        }
    }
}
