using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
namespace Replace_DocVariable_to_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../../Data/Template.docx"), FormatType.Docx))
            {
                WTable table = new WTable(document);
                table.ResetCells(1, 3);
                //Header FIRST ROW
                table[0, 0].Width = 13f;
                table[0, 1].Width = 128f;
                table[0, 1].AddParagraph().AppendText("Syncfusion");
                table[0, 2].Width = 80f;
                table[0, 2].AddParagraph().AppendText("DocIO");

                string docVariableToBereplaced = "InsertTable";
                //Finds the DocVariable field and replace it with a Table
                ReplaceDocVariableWithTable(document, docVariableToBereplaced, table);
                //Saves and closes the document instance
                document.Save("../../../Output/output.docx", FormatType.Docx);
            }
        }
        /// <summary>
        /// Finds the DocVariable field and replace it with a Table
        /// </summary>
        /// <param name="document">Word document in which you should replace the field with</param>
        /// <param name="docVariableToBeReplaced">The name of the Doc variable that you should replace</param>
        /// <param name="table">The table which will replace the Doc Variable</param>
        private static void ReplaceDocVariableWithTable(WordDocument document, string docVariableToBeReplaced, WTable table)
        {
            List<Entity> docFields = document.FindAllItemsByProperties(EntityType.Field, new string[2]{"FieldType", "FieldValue" }, new string[2] { FieldType.FieldDocVariable.ToString(), docVariableToBeReplaced });
            if(docFields != null)
            {
                foreach (Entity ent in docFields)
                {
                    WField docField = ent as WField;
                    WParagraph ownerPara = docField.OwnerParagraph;
                    int fieldIndex = ownerPara.ChildEntities.IndexOf(docField);
                    int textBodyIndex = ownerPara.OwnerTextBody.ChildEntities.IndexOf(ownerPara);
                    ownerPara.ChildEntities.Remove(docField);
                    //Split the paragraph and insert the table in the between
                    SplitAndInsertTable(ownerPara, table, fieldIndex, textBodyIndex);
                }
            }
        }
        /// <summary>
        /// Split the paragraph and insert the table in the between
        /// </summary>
        /// <param name="paragraph">Owner paragraph of the field</param>
        /// <param name="table">The table which will replace the Doc Variable</param>
        /// <param name="fieldIndex">Index of the paragraph item</param>
        /// <param name="textBodyIndex">Index of the text body</param>
        private static void SplitAndInsertTable(WParagraph paragraph, WTable table, int fieldIndex, int textBodyIndex)
        {
            WTextBody ownerTextBody = paragraph.OwnerTextBody;
            // Clone and trim the right-side paragraph (content after fieldIndex)
            WParagraph rightParagraph = (WParagraph)paragraph.Clone();
            for (int i = fieldIndex - 1; i >= 0; i--)
                rightParagraph.ChildEntities.RemoveAt(i);
            // Trim original paragraph to content before fieldIndex
            for (int i = paragraph.ChildEntities.Count - 1; i >= fieldIndex; i--)
                paragraph.ChildEntities.RemoveAt(i);
            bool removedParagraph = false;
            if (paragraph.ChildEntities.Count == 0)
            {
                ownerTextBody.ChildEntities.Remove(paragraph);
                removedParagraph = true;
            }
            int tableInsertIndex = removedParagraph ? textBodyIndex : textBodyIndex + 1;
            ownerTextBody.ChildEntities.Insert(tableInsertIndex, table);
            if (rightParagraph.ChildEntities.Count > 0)
                ownerTextBody.ChildEntities.Insert(tableInsertIndex + 1, rightParagraph);
        }
    }
}
