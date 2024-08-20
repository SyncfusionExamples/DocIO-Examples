using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_bookmark_start_and_end_around_textrange
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access paragraph in a Word document.
                    WParagraph paragraph = document.Sections[0].Paragraphs[3] as WParagraph;
                    //Create bookmarkstart and bookmarkend instance.
                    BookmarkStart bookmarkStart = new BookmarkStart(document, "Northwind");
                    BookmarkEnd bookmarkEnd = new BookmarkEnd(document, "Northwind");
                    //Add bookmarkstart at index zero.
                    paragraph.Items.Insert(0, bookmarkStart);
                    //Add bookmarkend at index 2.
                    paragraph.Items.Insert(2, bookmarkEnd);
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
