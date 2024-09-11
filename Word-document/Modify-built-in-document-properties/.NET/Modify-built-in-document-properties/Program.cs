using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;

namespace Modify_built_in_document_properties
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Accesses the built-in document properties.
                    Console.WriteLine("Title - {0}", document.BuiltinDocumentProperties.Title);
                    Console.WriteLine("Author - {0}", document.BuiltinDocumentProperties.Author);
                    //Modifies or sets the Built-in document properties.
                    document.BuiltinDocumentProperties.Author = "Andrew";
                    document.BuiltinDocumentProperties.LastAuthor = "Steven";
                    document.BuiltinDocumentProperties.CreateDate = new DateTime(1900, 12, 31, 12, 0, 0);
                    document.BuiltinDocumentProperties.LastSaveDate = new DateTime(1900, 12, 31, 12, 0, 0);
                    document.BuiltinDocumentProperties.LastPrinted = new DateTime(1900, 12, 31, 12, 0, 0);
                    document.BuiltinDocumentProperties.Title = "Sample Document";
                    document.BuiltinDocumentProperties.Subject = "Adventure Works Cycle";
                    document.BuiltinDocumentProperties.Category = "Technical documentation";
                    document.BuiltinDocumentProperties.Comments = "This is sample document.";
                    document.BuiltinDocumentProperties.RevisionNumber = "2";
                    document.BuiltinDocumentProperties.Company = "Adventure Works Cycle";
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
