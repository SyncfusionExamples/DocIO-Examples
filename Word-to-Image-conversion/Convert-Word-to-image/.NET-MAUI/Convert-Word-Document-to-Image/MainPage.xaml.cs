using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;
using System.Reflection;
using Convert_Word_Document_to_Image.SaveServices;

namespace Convert_Word_Document_to_Image;

public partial class MainPage : ContentPage
{
	int count = 0;

	public MainPage()
	{
		InitializeComponent();
	}

    private void ConvertWordtoImage(object sender, EventArgs e)
    {
        //Loading an existing Word document
        Assembly assembly = typeof(App).GetTypeInfo().Assembly;
        using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Convert_Word_Document_to_Image.Template.Input.docx"), FormatType.Docx))
        {
            //Instantiation of DocIORenderer for Word to Image conversion
            using (DocIORenderer render = new DocIORenderer())
            {
                //Convert the first page of the Word document into an image.
                Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
                //Reset the stream position.
                imageStream.Position = 0;
                //save and Launch the Image 
                SaveService saveService = new();
                saveService.SaveAndView("wordtoimage.jpeg", "application/jpeg", imageStream as MemoryStream);
            }
        }
    }
}

