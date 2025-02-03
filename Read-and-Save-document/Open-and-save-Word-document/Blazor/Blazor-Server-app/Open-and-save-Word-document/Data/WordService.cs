using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;


namespace Open_and_save_Word_document.Data
{
    public class WordService
    {
        public MemoryStream OpenAndSaveDocument()
        {
            using (FileStream sourceStreamPath = new FileStream(@"wwwroot/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    //Access the section in a Word document.
                    IWSection section = document.Sections[0];
                    //Add new paragraph to the section.
                    IWParagraph paragraph = section.AddParagraph();
                    paragraph.ParagraphFormat.FirstLineIndent = 36;
                    paragraph.BreakCharacterFormat.FontSize = 12f;
                    //Add new text to the paragraph.
                    IWTextRange textRange = paragraph.AppendText("In 2000, AdventureWorks Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.") as IWTextRange;
                    textRange.CharacterFormat.FontSize = 12f;

                    //Save the Word document to MemoryStream.
                    MemoryStream stream = new MemoryStream();
                    document.Save(stream, FormatType.Docx);
                    stream.Position = 0;
                    return stream;
                }
            }
        }
    }
}
