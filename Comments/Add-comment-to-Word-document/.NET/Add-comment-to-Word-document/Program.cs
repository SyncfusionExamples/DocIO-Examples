using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Add_comment_to_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds a section and a paragraph in the document.
                document.EnsureMinimal();
                IWParagraph paragraph = document.LastParagraph;
                //Appends text to the paragraph.
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Adds comment to a paragraph.
                WComment comment = paragraph.AppendComment("comment test");
                //Specifies the author of the comment.
                comment.Format.User = "Peter";
                //Specifies the initial of the author.
                comment.Format.UserInitials = "St";
                //Set the date and time for comment.
                comment.Format.DateTime = DateTime.Now;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
