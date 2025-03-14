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
                    // Update checkbox form fields
                    List<Entity> checkBoxes = document.FindAllItemsByProperty(EntityType.CheckBox, null, null);
                    foreach (Entity entity in checkBoxes)
                    {
                        WCheckBox checkBox = (WCheckBox)entity;
                        checkBox.SizeType = CheckBoxSizeType.Exactly;
                        checkBox.CheckBoxSize = 20;
                    }

                    // Update dropdown form fields
                    List<Entity> dropDowns = document.FindAllItemsByProperty(EntityType.DropDownFormField, null, null);
                    foreach (Entity entity in dropDowns)
                    {
                        SetFontSizeForFormField((WDropDownFormField)entity);
                    }

                    // Update text form fields
                    List<Entity> textFormFields = document.FindAllItemsByProperty(EntityType.TextFormField, null, null);
                    foreach (Entity entity in textFormFields)
                    {
                        SetFontSizeForFormField((WTextFormField)entity);
                    }
                    // Save the modified document
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
			/// <summary>
            /// Sets the font size for text ranges within a form field until the Field End marker is reached.
            /// </summary>
            static void SetFontSizeForFormField(Entity formField)
            {
                Entity currentEntity = formField;
                while (currentEntity.NextSibling != null)
                {
                    if (currentEntity is WTextRange textRange)
                    {
                        textRange.CharacterFormat.FontSize = 20;
                    }
                    else if (currentEntity is WFieldMark fieldMark && fieldMark.Type == FieldMarkType.FieldEnd)
                    {
                        break;
                    }
                    currentEntity = (Entity)currentEntity.NextSibling;
                }
            }
        }
    }
}
