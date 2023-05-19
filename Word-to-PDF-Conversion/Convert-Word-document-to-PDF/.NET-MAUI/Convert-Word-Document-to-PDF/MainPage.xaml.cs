
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Reflection;
using Convert_Word_Document_to_PDF.SaveServices;

namespace Convert_Word_Document_to_PDF;

public partial class MainPage : ContentPage
{
	int count = 0;

	public MainPage()
	{
		InitializeComponent();
	}

	private void ConvertWordtoPDF(object sender, EventArgs e)
	{
        //Loading an existing Word document
        Assembly assembly = typeof(App).GetTypeInfo().Assembly;      
        using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Convert_Word_Document_to_PDF.Assets.Input.docx"), FormatType.Docx))
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

                    //save and Launch the PDF document
                    SaveService saveService = new();                 
                    saveService.SaveAndView("Sample.pdf", "application/pdf", stream);
                }
            }
        }
    }
}

