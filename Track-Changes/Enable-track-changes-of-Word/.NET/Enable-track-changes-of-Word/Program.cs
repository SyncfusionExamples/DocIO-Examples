using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Enable_track_changes_of_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Appends text to the paragraph.
                IWTextRange text = paragraph.AppendText("This sample illustrates how to track the changes made to the word document. ");
                //Sets font name and size for text.
                text.CharacterFormat.FontName = "Times New Roman";
                text.CharacterFormat.FontSize = 14;
                text = paragraph.AppendText("This track changes is useful in shared environment.");
                text.CharacterFormat.FontSize = 12;
                //Turns on the track changes option.
                document.TrackChanges = true;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
