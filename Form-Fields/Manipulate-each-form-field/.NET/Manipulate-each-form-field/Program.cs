using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;

// Opens template Word document
using (WordDocument document = new WordDocument(@"../../../Data/input.docx", FormatType.Docx))
{
    // Find and update checkbox form fields
    List<Entity> checkBoxes = document.FindAllItemsByProperty(EntityType.CheckBox, null, null);
    foreach (Entity entity in checkBoxes)
    {
        WCheckBox checkBox = entity as WCheckBox;
        checkBox.SizeType = CheckBoxSizeType.Exactly;
        checkBox.CheckBoxSize = 20;
    }

    // Find and update dropdown form fields
    List<Entity> dropDowns = document.FindAllItemsByProperty(EntityType.DropDownFormField, null, null);
    foreach (Entity entity in dropDowns)
    {
        WDropDownFormField dropdown = entity as WDropDownFormField;
        //Modifies the dropdown items
        dropdown.DropDownItems.Remove(1);
        dropdown.DropDownSelectedIndex = 0;
        dropdown.CharacterFormat.FontName = "Arial";
    }

    // Find and update text form fields
    List<Entity> textFormFields = document.FindAllItemsByProperty(EntityType.TextFormField, null, null);
    foreach (Entity entity in textFormFields)
    {
        WTextFormField textField = entity as WTextFormField;
        //Modifies the text form field
        textField.Type = TextFormFieldType.RegularText;
        textField.StringFormat = "";
        textField.DefaultText = "Default text";
        textField.Text = "Default text";
        textField.CalculateOnExit = false;
    }
    //Save and dispose document
    document.Save(@"../../../Output/Result.docx", FormatType.Docx);
}