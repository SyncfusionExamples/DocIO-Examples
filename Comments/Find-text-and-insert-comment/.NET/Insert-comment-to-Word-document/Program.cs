using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Insert_comment_to_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Find all occurrence of a particular text ending with comma in the document using regex.
                    TextSelection[] textSelection = document.FindAll(new Regex("\\w+,"));
                    if (textSelection != null)
                    {
                        //Iterates through each occurrence and comment it.
                        for (int i = 0; i < textSelection.Count(); i++)
                        {
                            //Get the found text as a single text range.
                            WTextRange textRange = textSelection[i].GetAsOneRange();
                            //Get the owner paragraph of the found text.
                            WParagraph paragraph = textRange.OwnerParagraph;
                            //Get the index of the found text.
                            int textIndex = paragraph.ChildEntities.IndexOf(textRange);
                            //Add comment to a paragraph.
                            WComment comment = paragraph.AppendComment("comment test_" + i);
                            //Specify the author of the comment.
                            comment.Format.User = "Peter";
                            //Specify the initial of the author.
                            comment.Format.UserInitials = "St";
                            //Set the date and time for the comment.
                            comment.Format.DateTime = DateTime.Now;
                            //Insert the comment next to the textrange.
                            paragraph.ChildEntities.Insert(textIndex + 1, comment);
                            //Add the paragraph items to the commented items.
                            comment.AddCommentedItem(textRange);
                        }
                    }
                    //Create the file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
