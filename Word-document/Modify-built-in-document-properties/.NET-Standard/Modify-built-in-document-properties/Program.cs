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
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Accesses the built-in document properties.
                    Console.WriteLine("Title - {0}", document.BuiltinDocumentProperties.Title);
                    Console.WriteLine("Author - {0}", document.BuiltinDocumentProperties.Author);
                    //Modifies or sets the category and company Built-in document properties.
                    document.BuiltinDocumentProperties.Category = "Sales reports";
                    document.BuiltinDocumentProperties.Company = "Northwind traders";
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
