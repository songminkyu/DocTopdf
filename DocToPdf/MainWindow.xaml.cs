using DocToPdf.ViewModel;
using log4net;
using log4net.Config;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace DocToPdf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {        
        public MainWindow()
        {            
            InitializeComponent();
            DataContext = new DocToConvViewModel();
        }
    }
}