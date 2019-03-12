using Microsoft.Win32;
using System.Windows;
using System.IO;

namespace DocxToWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                DocxToFlowDocumentConverter converter = new DocxToFlowDocumentConverter(new FileStream(openFileDialog.FileName, FileMode.Open));
                converter.Read();
                documentViewer.Document = converter.Document;
            }
        }
    }
}
