using Syncfusion.DocIO.DLS;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.ComponentModel;
using System.Windows;


namespace Convert_Word_Document_to_PDF
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
        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Automatic))
            {
                //Instantiation of DocToPDFConverter for Word to PDF conversion
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    //Converts Word document into PDF document
                    using (PdfDocument pdfDocument = converter.ConvertToPDF(document))
                    {
                        //Saves the PDF document
                        pdfDocument.Save(Path.GetFullPath(@"../../Sample.pdf"));
                    }
                };
            }
            MessageBox.Show("PDF converted successfully");
        }                    
    }
}
