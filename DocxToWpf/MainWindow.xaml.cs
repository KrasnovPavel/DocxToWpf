// This is an independent project of an individual developer. Dear PVS-Studio, please check it.
// PVS-Studio Static Code Analyzer for C, C++ and C#: http://www.viva64.com

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

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
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
