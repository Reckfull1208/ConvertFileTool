using Microsoft.Win32;
using Spire.Pdf;
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

namespace ConverterFileTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string BasePath = string.Empty;
        public MainWindow()
        {
            InitializeComponent();
            BasePath = AppDomain.CurrentDomain.BaseDirectory;
        }

        private void PdfToWord_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Text Documents (.pdf)|*.pdf|All files (*.*) | *.*";
            var result = openFile.ShowDialog();
            if(result == true)
            {
                var path = openFile.FileName;
                var fileName = openFile.SafeFileName;
                Spire.Pdf.PdfDocument pdfDoc = new Spire.Pdf.PdfDocument();
                pdfDoc.LoadFromFile(path);  
                
                pdfDoc.ConvertOptions.SetPdfToDocOptions();
                pdfDoc.SaveToFile(System.IO.Path.Combine(BasePath, $"{fileName}.doc"), Spire.Pdf.FileFormat.DOC); 
                pdfDoc.Close();
                pdfDoc.Dispose();
            }
        }
    }
}
