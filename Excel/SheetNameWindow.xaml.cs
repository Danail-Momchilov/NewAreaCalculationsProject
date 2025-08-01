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

namespace AreaCalculations
{
    /// <summary>
    /// Interaction logic for SheetNameWindow.xaml
    /// </summary>
    public partial class SheetNameWindow : Window
    {
        private string sheetInputName { get; set; }
        public SheetNameWindow()
        {
            InitializeComponent();
        }
        private void inputButtonClick(object sender, RoutedEventArgs e)
        {
            try
            {
                this.sheetInputName = sheetName.Text;
            }
            catch { }

            this.Close();
        }
    }
}
