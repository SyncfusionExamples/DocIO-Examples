using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using static System.Collections.Specialized.BitVector32;
using System.Reflection.Metadata;

namespace Merge_field_inside_IF_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            WordDocument document = new WordDocument();
            //Add new section to the document.
            IWSection section = document.AddSection();
            //Add new paragraph to section.
            WParagraph paragraph = section.AddParagraph() as WParagraph;
            
            //Create a new instance of IF field.
            WField field = paragraph.AppendField("If", FieldType.FieldIf) as WField;
            //Specifies the field code.
            InsertIfFieldCode(paragraph, field);

            //Execute Mail merge.
            string[] fieldName = { "Gender", "Male", "Female" };
            string[] fieldValue = { "M", "Mr.Andrew", "Miss.Nancy" };
            document.MailMerge.Execute(fieldName, fieldValue);
            //Update the fields.
            document.UpdateDocumentFields();

            //Save and close the document.
            FileStream outputStream = new FileStream("../../../Sample.docx", FileMode.Create, FileAccess.Write);
            document.Save(outputStream, FormatType.Docx);
            document.Close();
        }
        /// <summary>
        /// Insert the field code with nested field 
        /// </summary>
        private static void InsertIfFieldCode(WParagraph paragraph, WField field)
        {
            //Insert the field code based on IF field syntax.
            //IF field syntax - { IF Expression1OperatorExpression2TrueTextFalseText} 

            //Get the index of the If field.
            int fieldIndex = paragraph.Items.IndexOf(field) + 1;
            //Add the field code.
            field.FieldCode = "IF ";
            //To insert Merge field after "IF" field code increment the index.
            fieldIndex++;
            InsertText("\"", ref fieldIndex, paragraph);
            InsertMergeField("Gender", ref fieldIndex, paragraph);
            InsertText("\" = \"", ref fieldIndex, paragraph);
            InsertMergeField("Gender", ref fieldIndex, paragraph);
            InsertText("\" \"", ref fieldIndex, paragraph);
            InsertMergeField("Male", ref fieldIndex, paragraph);
            InsertText("\" \"", ref fieldIndex, paragraph);
            InsertMergeField("Female", ref fieldIndex, paragraph);
            InsertText("\"", ref fieldIndex, paragraph);
        }
        /// <summary>
        /// Insert text such as quote, operator.
        /// </summary>
        private static void InsertText( string text,ref int fieldIndex, WParagraph paragraph)
        {
            //Insert the Operator in a textrange.
            WTextRange textRange = new WTextRange(paragraph.Document);
            textRange.Text = text;
            //Insert the textrange as field code item.
            paragraph.Items.Insert(fieldIndex, textRange);
            fieldIndex++;
        }
        /// <summary>
        /// Insert merge field at the given index
        /// </summary>
        private static void InsertMergeField(string fieldName, ref int fieldIndex, WParagraph paragraph)
        {
            WParagraph para = new WParagraph(paragraph.Document);
            para.AppendField(fieldName, FieldType.FieldMergeField);
            int count = para.ChildEntities.Count;
            //As the child entity is a field, if we insert the field it automaticlly inserts the complete field structure.
            paragraph.ChildEntities.Insert(fieldIndex, para.ChildEntities[0]);
            fieldIndex += count;
        }
    }
}
