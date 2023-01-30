using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;


namespace Modify_built_in_document_properties.Data
{
    public class WordService
    {
        public MemoryStream CreateWord()
        {
            using (FileStream sourceStreamPath = new FileStream(@"wwwroot/data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opening a document.
                using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    //Accesses the built-in document properties
                    Console.WriteLine("Title - {0}", document.BuiltinDocumentProperties.Title);
                    Console.WriteLine("Author - {0}", document.BuiltinDocumentProperties.Author);
                    //Modifies or sets the category and company Built-in document properties
                    document.BuiltinDocumentProperties.Category = "Sales reports";
                    document.BuiltinDocumentProperties.Company = "Northwind traders";

                    //Saves the Word document to MemoryStream.
                    using (MemoryStream stream = new MemoryStream())
                    {
                        document.Save(stream, FormatType.Docx);
                        stream.Position = 0;
                        return stream;
                    }
                }
            }
        }
    }
}
