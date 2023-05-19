using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Reflection;
using System.IO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Convert_Word_Document_to_PDF
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void OnButtonClicked(object sender, EventArgs e)
        {
            //Loading an existing Word document
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;

            using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Convert-Word-Document-to-PDF.Assets.Input.docx"), FormatType.Docx))
            {
                //Instantiation of DocIORenderer for Word to PDF conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Converts Word document into PDF document
                    using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                    {
                        //Saves the PDF document to MemoryStream.
                        MemoryStream stream = new MemoryStream();
                        pdfDocument.Save(stream);

                        //Save the stream as a file in the device and invoke it for viewing.
                        Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("Sample.pdf", "application/pdf", stream);
                    }
                }               
            }                         
        }       
    }
}
