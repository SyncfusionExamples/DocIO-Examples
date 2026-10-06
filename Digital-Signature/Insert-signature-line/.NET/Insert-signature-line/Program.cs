using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Insert_signature_line
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Template.docx"));
            //Gets the first section of the document.
            IWSection section = document.Sections[0];
            //Adds new paragraph to the section.
            IWParagraph paragraph = section.AddParagraph();
            //Adds new text to the paragraph
            paragraph.AppendText("Please sign below: ");
            //Adds a new paragraph that will host the signature line.
            IWParagraph signatureParagraph = section.AddParagraph();
            //Configures the signature line settings.
            SignatureLineSettings settings = new SignatureLineSettings();
            settings.Signer = "John Doe";
            settings.SignerTitle = "Manager";
            settings.Email = "john.doe@example.com";
            settings.Instructions = "Please review and sign.";
            //Specifies whether the signer can attach comments when signing.
            settings.AllowComments = true;
            //Specifies whether the sign date is displayed on the signature line.
            settings.ShowDate = true;
            //Specifies whether the default signing instructions are used (false uses the custom Instructions value).
            settings.DefaultInstructions = false;
            //Inserts the signature line into the new paragraph with the specified dimensions.
            IWPicture picture = signatureParagraph.AppendSignatureLine(settings, 200, 100);
            //Saves the Word document to file.
            document.Save(Path.GetFullPath(@"Output\Result.docx"), FormatType.Docx);
            //Closes the document
            document.Close();
        }
    }
}

