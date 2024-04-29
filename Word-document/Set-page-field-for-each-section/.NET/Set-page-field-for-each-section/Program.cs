using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Set_page_field_for_each_section
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance of WordDocument.
            using (WordDocument document = new WordDocument())
            {

                //Adds new section to the document.
                IWSection section1 = document.AddSection();
                //Inserts the default page header.
                IWParagraph paragraph = section1.HeadersFooters.OddHeader.AddParagraph();
                //Adds the new Page field in header with field name and its type
                IWField field1 = paragraph.AppendField("Page", FieldType.FieldPage);

                //Adds a paragraph to created section.
                paragraph = section1.AddParagraph();
                //Appends the text to the created paragraph.
                paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

                //Sets the section break.
                section1.BreakCode = SectionBreakCode.NewPage;

                IWSection section2 = document.AddSection();
                //Inserts the default page header.
                paragraph = section2.HeadersFooters.OddHeader.AddParagraph();
                //Adds the new Page field in header with field name and its type
                IWField field2 = paragraph.AppendField("Page", FieldType.FieldPage);


                //Updates the fields present in a document, to update page fields.
                document.UpdateDocumentFields(true);

                //Find and print the page numbers in the Word document.
                FindAndPrintPageNumbers(document);

                //Saves the Word document to file system.    
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    document.Save(outputStream,FormatType.Docx);
                }
            }

        }
        /// <summary>
        /// Finds and prints the page numbers in a Word document.
        /// </summary>
        private static void FindAndPrintPageNumbers(WordDocument document)
        {
            //Find all page fields by EntityType in Word document.
            List<Entity> pageField = document.FindAllItemsByProperty(EntityType.Field, "FieldType", "FieldPage");

            if (pageField != null)
            {
                //Print the page numbers in the Word document.
                for (int i = 0; i < pageField.Count; i++)
                {
                    Console.WriteLine((pageField[i] as WField).Text);
                }
            }
        }
    }
}
