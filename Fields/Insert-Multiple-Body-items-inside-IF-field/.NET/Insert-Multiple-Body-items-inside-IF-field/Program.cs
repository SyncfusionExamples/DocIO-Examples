using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Insert_Multiple_Body_items_inside_IF_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add new section to the document.
                IWSection section = document.AddSection();
                //Add new paragraph to section.
                WParagraph paragraph = section.AddParagraph() as WParagraph;

                //Create a new instance of IF field.
                WField field = paragraph.AppendField("If", FieldType.FieldIf) as WField;
                //Specifies the field code.
                InsertIfFieldCode(paragraph, field);

                //Update the fields.
                document.UpdateDocumentFields();

                //Save and close the document.
                using (FileStream outputStream = new FileStream("Result.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
                document.Close();
            }
        }
        /// <summary>
        /// Modify the field code.
        /// </summary>
        private static void InsertIfFieldCode(WParagraph paragraph, WField field)
        {
            //Get the index of the If field i.e., Field code index.
            int fieldIndex = paragraph.Items.IndexOf(field) + 1;
            //Add the field code.
            field.FieldCode = "IF \"M\" = \"M\" ";
            //Move the field separator to field end to temporary paragraph.
            WParagraph lastPara = new WParagraph(paragraph.Document); 
            MoveFieldMark(paragraph, fieldIndex + 1, lastPara);
            //Set the true statement
            paragraph = InsertTrueStatement(paragraph);
            //Set the false statement
            paragraph = InsertFalseStatement(paragraph);
            //Move the content from temporary paragraph to last paragraph.
            MoveFieldMark(lastPara,0, paragraph);
        }
        /// <summary>
        /// Move the field items to another paragraph.
        /// </summary>
        private static void MoveFieldMark(WParagraph paragraph, int fieldIndex, WParagraph lastPara)
        {
            //Move the field separator to field end to the last paragraph.
            for(int i = fieldIndex; i < paragraph.Items.Count;)
                lastPara.Items.Add(paragraph.Items[i]);
        }
        /// <summary>
        /// Insert the multiple body items as true statement.
        /// </summary>
        private static WParagraph InsertTrueStatement(WParagraph paragraph)
        {
            WTextBody ownerTextBody = paragraph.OwnerTextBody;
            //Add text to the existing paragraph.
            paragraph.AppendText("\"List of male candidates:");
            //Add a table.
            WTable table = ownerTextBody.AddTable() as WTable;
            //Add first row.
            WTableRow row = table.AddRow() as WTableRow;
            row.AddCell().AddParagraph().AppendText("1.");
            row.AddCell().AddParagraph().AppendText("Mr. Peter");
            //Add second row.
            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("2.");
            row.Cells[1].AddParagraph().AppendText("Mr. Andrew");
            //Add third row.
            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("3.");
            row.Cells[1].AddParagraph().AppendText("Mr. Thomas");
            //Add a new paragraph for true statement end quote.
            WParagraph lastPara = ownerTextBody.AddParagraph() as WParagraph;
            lastPara.AppendText("\" ");
            return lastPara;
        }
        /// <summary>
        /// Insert multiple body items as false statement.
        /// </summary>
        private static WParagraph InsertFalseStatement(WParagraph paragraph)
        {
            WTextBody ownerTextBody = paragraph.OwnerTextBody;
            //Add text to the existing paragraph.
            paragraph.AppendText("\"List of female candidates:");
            //Add a table.
            WTable table = ownerTextBody.AddTable() as WTable;
            //Add first row.
            WTableRow row = table.AddRow() as WTableRow;
            row.AddCell().AddParagraph().AppendText("1.");
            row.AddCell().AddParagraph().AppendText("Mrs. Nancy");
            //Add second row.
            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("2.");
            row.Cells[1].AddParagraph().AppendText("Miss. Janet");
            //Add third row.
            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("3.");
            row.Cells[1].AddParagraph().AppendText("Mrs. Margaret");
            //Add a new paragraph for true statement end quote.
            WParagraph lastPara = ownerTextBody.AddParagraph() as WParagraph;
            lastPara.AppendText("\"");
            return lastPara;
        }
    }
}