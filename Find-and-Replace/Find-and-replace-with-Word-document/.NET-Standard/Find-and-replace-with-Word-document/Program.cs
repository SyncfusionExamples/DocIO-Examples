using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace Find_and_replace_with_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"D:\Support\549924\Template1.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    using (FileStream sourceFileStream = new FileStream(Path.GetFullPath(@"D:\Support\549924\section3.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source Word document to copy all the content. 
                        using (WordDocument sourceWordDocument = new WordDocument(sourceFileStream, FormatType.Automatic))
                        {
                            //Get all the content as Word document part.
                            WordDocumentPart wordDocumentPart = new WordDocumentPart(sourceWordDocument);
                            //Create the bookmark navigator instance to access the bookmark.
                            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                            //Move the virtual cursor to the location before the end of the bookmark "Adventure_Bkmk".
                            bookmarkNavigator.MoveToBookmark("Key_Responsibilities");
                            document.ImportOptions = ImportOptions.MergeFormatting;
                            //Replace the bookmark content with Word document part.
                            bookmarkNavigator.ReplaceContent(wordDocumentPart);
                            //Create file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the Word document to file stream. 
                                document.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                    ////Finds all the content placeholder text in the Word document.
                    //TextSelection[] textSelections = document.FindAll(new Regex(@"\<<(.*)\>>"));
                    //for (int i = 0; i < textSelections.Length; i++)
                    //{
                    //    //Replaces the content placeholder text with desired Word document.
                    //    using (WordDocument subDocument = new WordDocument(new FileStream(Path.GetFullPath(@"D:\Support\549924\section3.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite), FormatType.Docx))
                    //    {
                    //        document.Replace(textSelections[i].SelectedText, subDocument, true, true);
                    //    }
                    //}
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"D:\Support\549924\Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
