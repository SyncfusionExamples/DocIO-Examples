using Open_and_save_Word_document.SaveServices;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Reflection;

namespace Open_and_save_Word_document;

public partial class MainPage : ContentPage
{
	int count = 0;

	public MainPage()
	{
		InitializeComponent();
	}

	private void OpenAndSaveDocument(object sender, EventArgs e)
	{
        //Load an existing Word document.
        Assembly assembly = typeof(App).GetTypeInfo().Assembly;
        using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Open_and_save_Word_document.Resources.Input.docx"),FormatType.Docx))
        {
            //Access the section in a Word document.
            IWSection section = document.Sections[0];
            //Add a new paragraph to the section.
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.FirstLineIndent = 36;
            paragraph.BreakCharacterFormat.FontSize = 12f;
            IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
            text.CharacterFormat.FontSize = 12f;
            //Saves the Word document to the memory stream.
            using MemoryStream ms = new();
            document.Save(ms, FormatType.Docx);
            ms.Position = 0;
            //Save the memory stream as a file.
            SaveService saveService = new();
            saveService.SaveAndView("Sample.docx", "application/msword", ms);
        }
    }
}

