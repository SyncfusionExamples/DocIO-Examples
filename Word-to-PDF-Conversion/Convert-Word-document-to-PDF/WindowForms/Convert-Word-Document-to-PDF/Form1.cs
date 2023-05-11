using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Convert_Word_Document_to_PDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnConvert_Click(object sender, EventArgs e)
        {
            //Load the existing Word document 
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
