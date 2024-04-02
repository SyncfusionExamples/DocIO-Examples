using System.IO;
using System.Windows;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Convert_Word_to_EPUB
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
        private void btnWordToEPUB_Click(object sender, RoutedEventArgs e)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../../Data/Input.docx"), FormatType.Automatic))
            {
                //Exports the fonts used in the document
                document.SaveOptions.EPubExportFont = true;
                //Exports header and footer
                document.SaveOptions.HtmlExportHeadersFooters = true;

                //Save the Word document.
                document.Save(@"../../../WordToEPUB.epub", FormatType.EPub);
            }
            MessageBox.Show("Word document generated successfully");
        }
    }
}