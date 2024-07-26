using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_an_existing_paragraph
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an input Word template.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Gets the text body of first section.
                    WTextBody textBody = document.Sections[0].Body;
                    //Gets the paragraph at index 1.
                    WParagraph paragraph = textBody.Paragraphs[1];
                    //Iterates through the child elements of paragraph.
                    foreach (ParagraphItem item in paragraph.ChildEntities)
                    {
                        if (item is WTextRange)
                        {
                            WTextRange text = item as WTextRange;
                            //Modifies the character format of the text.
                            text.CharacterFormat.Bold = true;
                            break;
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
