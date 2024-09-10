using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace RTL_paragraph
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Gets the text body of first section.
                    WTextBody textBody = document.Sections[0].Body;
                    //Gets the paragraph at index 1.
                    WParagraph paragraph = textBody.Paragraphs[1];
                    //Gets a value indicating whether the paragraph is right-to-left. True indicates the paragraph direction is RTL.
                    bool isRTL = paragraph.ParagraphFormat.Bidi;
                    //Sets RTL direction for a paragraph.
                    if (!isRTL)
                        paragraph.ParagraphFormat.Bidi = true;
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
}
