using System;
using System.Collections.Generic;
using System.IO;
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
using Aspose.Finance.Xbrl;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlbr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string sourceDir = Directory.GetCurrentDirectory();
        
        public MainWindow()
        {
            InitializeComponent();
            ruta.Text = sourceDir.Substring(0, sourceDir.IndexOf("bin"));
            convertToXLSX(sourceDir.Substring(0, sourceDir.IndexOf("bin")));
        }

        public void convertToXLSX(string path) {
            // Load input XBRL file
            XbrlDocument document = new XbrlDocument(path + @"data\deposito.xbrl");

            // Set SaveOptions for output file
            SaveOptions saveOptions = new SaveOptions();
            saveOptions.SaveFormat = SaveFormat.XLSX;

            // Convert XBRL file to XLSX Excel Worksheet format
            document.Save(path + @"data\deposito.xlsx", saveOptions);
        }
    }
}