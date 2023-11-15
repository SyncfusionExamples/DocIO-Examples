
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Open_and_save_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    //Access the section in a Word document.
                    IWSection section = document.Sections[0];
                    //Add a new paragraph to the section.
                    IWParagraph paragraph = section.AddParagraph();
                    paragraph.ParagraphFormat.FirstLineIndent = 36;
                    paragraph.BreakCharacterFormat.FontSize = 12f;
                    IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
                    text.CharacterFormat.FontSize = 12f;
                    //Create a FileStream to save the Word file.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                    {
                        //Save the Word file.
                        document.Save(outputStream, FormatType.Docx);
                    }
                        
                }
            }
        }
    }
}
