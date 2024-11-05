using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (WordDocument document = new WordDocument())
{
    // Add one section and one paragraph to the document.
    document.EnsureMinimal();

    // Create a CheckBoxState for the checked state, using a tick symbol in the Wingdings font
    CheckBoxState tickState = new CheckBoxState
    {
        Font = "Wingdings",
        Value = "\uF0FE" // Unicode for the tick symbol (✓) in Wingdings
    };
    // Create a CheckBoxState for the unchecked state, using an empty box symbol in the Wingdings font
    CheckBoxState unTickState = new CheckBoxState
    {
        Font = "Wingdings",
        Value = "\uF0A8" // Unicode for the empty box symbol in Wingdings
    };

    // Gets the last paragraph.
    WParagraph paragraph = document.LastParagraph;
    // Add text to the paragraph.
    paragraph.AppendText("Gender:\tFemale ");
    // Append checkbox content control to the paragraph  for the "checked" option.
    IInlineContentControl checkBox = paragraph.AppendInlineContentControl(ContentControlType.CheckBox);
    // Set the checked state of the checkbox content control to display the tick symbol when selected
    checkBox.ContentControlProperties.CheckedState = tickState;
    // Set the unchecked state of the checkbox content control to display an empty box when not selected
    checkBox.ContentControlProperties.UncheckedState = unTickState;
    // Set the initial state of the "Female" checkbox to checked
    checkBox.ContentControlProperties.IsChecked = true;

    // Add text to the paragraph.
    paragraph.AppendText("\tMale ");
    // Append checkbox content control to the paragraph  for the "unchecked" option.
    checkBox = paragraph.AppendInlineContentControl(ContentControlType.CheckBox);
    // Set the checked and unchecked states.
    checkBox.ContentControlProperties.CheckedState = tickState;
    checkBox.ContentControlProperties.UncheckedState = unTickState;
    // Set the initial state of the "Male" checkbox to unchecked
    checkBox.ContentControlProperties.IsChecked = false;

    // Save the document.
    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
    {
        document.Save(outputStream, FormatType.Docx);
    }
}
