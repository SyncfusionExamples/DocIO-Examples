using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_bookmark_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a new section and paragraph in the document. 
                document.EnsureMinimal();
                //Get the last paragraph. 
                IWParagraph paragraph = document.LastParagraph;
                //Adds a new bookmark start into paragraph with name "Northwind".
                paragraph.AppendBookmarkStart("Northwind");
                //Adds a text between the bookmark start and end into paragraph.
                paragraph.AppendText("The Northwind sample database (Northwind.mdb) is included with all versions of Access. It provides data you can experiment with and database objects that demonstrate features you might want to implement in your own databases.");
                //Adds a new bookmark end into paragraph with name "Northwind".
                paragraph.AppendBookmarkEnd("Northwind");
                //Adds a text after the bookmark end.
                paragraph.AppendText(" Using Northwind, you can become familiar with how a relational database is structured and how the database objects work together to help you enter, store, manipulate, and print your data.");
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
