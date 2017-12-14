using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace demo_plugin.forms
{
    /// <summary>
    /// Interaction logic for DemoPane.xaml
    /// </summary>
    public partial class DemoPane : UserControl
    {
        public DemoPane()
        {
            InitializeComponent();
        }

        private void BtnOkClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("You clicked OK.");
        }

        private void BtnCancelClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("You clicked Cancel.");
        }
    }
}
