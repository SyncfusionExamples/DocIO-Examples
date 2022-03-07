using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Apply_number_format_for_SEQ_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = CreateDocument())
            {
                //Accesses sequence field in the document.
                WSeqField field = (document.LastSection.Body.ChildEntities[0] as WParagraph).ChildEntities[0] as WSeqField;
                //Applies the number format for sequence field.
                field.NumberFormat = CaptionNumberingFormat.Roman;
                //Accesses sequence field in the document.
                field = (document.LastSection.Body.ChildEntities[1] as WParagraph).ChildEntities[0] as WSeqField;
                //Applies the number format for sequence field.
                field.NumberFormat = CaptionNumberingFormat.Roman;
                //Accesses sequence field in the document.
                field = (document.LastSection.Body.ChildEntities[2] as WParagraph).ChildEntities[0] as WSeqField;
                //Applies the number format for sequence field.
                field.NumberFormat = CaptionNumberingFormat.Roman;
                //Updates the document fields.
                document.UpdateDocumentFields();
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        private static WordDocument CreateDocument()
        {
            //Creates a new document.
            WordDocument document = new WordDocument();
            //Adds a new section to the document.
            IWSection section = document.AddSection();
            //Sets margin of the section.
            section.PageSetup.Margins.All = 72;
            //Adds a paragraph to the section.
            IWParagraph paragraph = section.AddParagraph();
            paragraph.AppendField("List", FieldType.FieldSequence);
            paragraph.AppendText(".Item1");
            //Adds a paragraph to the section.
            paragraph = section.AddParagraph();
            paragraph.AppendField("List", FieldType.FieldSequence);
            paragraph.AppendText(".Item2");
            //Adds a paragraph to the section.
            paragraph = section.AddParagraph();
            paragraph.AppendField("List", FieldType.FieldSequence);
            paragraph.AppendText(".Item3");
            return document;
        }
    }
}
