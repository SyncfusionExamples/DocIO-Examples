using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Retrieve_commented_word
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load the existing Word document.
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                //Iterate the comments in the Word document.
                foreach (WComment comment in document.Comments)
                {
                    //Get the commented word or part of a particular comment.
                    if (comment.TextBody.LastParagraph.Text == "This is the second comment.")
                    {
                        ParagraphItemCollection paragraphItem = comment.CommentedItems;
                    }

                }
                //Save the Word document to file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
            
        }
    }
}
