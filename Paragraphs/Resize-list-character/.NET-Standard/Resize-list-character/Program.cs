using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Resize_list_character
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Get the textbody content and adds it to document section.
                    WTextBody textbody = document.Sections[0].Body;
                    //Iterate through thedocument paragraphs.
                    foreach (WParagraph paragraph in textbody.Paragraphs)
                    {
                        //Get the symbol from the paragraph items.
                        foreach (ParagraphItem item in paragraph.ChildEntities)
                        {
                            //Change the list character size.
                            if (paragraph.ListFormat != null && paragraph.ListFormat.CurrentListLevel != null)
                                paragraph.ListFormat.CurrentListLevel.CharacterFormat.FontSize = 25;
                        }
                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
