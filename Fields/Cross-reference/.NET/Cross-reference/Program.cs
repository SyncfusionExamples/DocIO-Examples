using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Cross_reference
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of a WordDocument.
            using (WordDocument document = new WordDocument())
            {
                //Adds a new section into the Word Document.
                IWSection section = document.AddSection();
                //Adds a new paragraph into Word document.
                IWParagraph paragraph = section.AddParagraph();
                //Adds text, bookmark start and end in the paragraph.
                paragraph.AppendBookmarkStart("Title");
                paragraph.AppendText("Northwind Database");
                paragraph.AppendBookmarkEnd("Title");
                paragraph = section.AddParagraph();
                paragraph.AppendText("The Northwind sample database (Northwind.mdb) is included with all versions of Access. It provides data you can experiment with and database objects that demonstrate features you might want to implement in your own databases.");
                section = document.AddSection();
                section.AddParagraph();
                paragraph = section.AddParagraph() as WParagraph;
                //Gets the collection of bookmark start in the word document.
                List<Entity> items = document.GetCrossReferenceItems(ReferenceType.Bookmark);
                paragraph.AppendText("Bookmark Cross Reference starts here ");
                //Appends the cross reference for bookmark “Title” with ContentText as reference kind.
                paragraph.AppendCrossReference(ReferenceType.Bookmark, ReferenceKind.ContentText, items[0], true, false, false, string.Empty);
                //Updates the document Fields.
                document.UpdateDocumentFields();
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
