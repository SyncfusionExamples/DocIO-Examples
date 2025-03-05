using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace SpecifyFormFieldSize
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the template Word document.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    // Find and update size for all checkbox form fields.
                    List<Entity> checkBoxes = document.FindAllItemsByProperty(EntityType.CheckBox, null, null);
                    foreach (Entity entity in checkBoxes)
                    {
                        WCheckBox checkBox = (WCheckBox)entity;
                        checkBox.SizeType = CheckBoxSizeType.Exactly;
                        checkBox.CheckBoxSize = 20; 
                    }

                    // Find and update size for all text form fields.
                    List<Entity> textFormFields = document.FindAllItemsByProperty(EntityType.TextFormField, null, null);
                    foreach (Entity entity in textFormFields)
                    {
                        WTextFormField textFormField = (WTextFormField)entity;
                        Entity currentEntity = textFormField;

                        // Iterate through sibling items until reaching the Field End marker.
                        while (currentEntity.NextSibling != null)
                        {
                            if (currentEntity is WTextRange)
                            {
                                // Set font size for text ranges within the form field.
                                (currentEntity as WTextRange).CharacterFormat.FontSize = 14;
                            }
                            else if (currentEntity is WFieldMark fieldMark && fieldMark.Type == FieldMarkType.FieldEnd)
                            {
                                break;
                            }
                            // Move to the next sibling entity.
                            currentEntity = (Entity)currentEntity.NextSibling;
                        }
                    }
                    // Save the modified document
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
