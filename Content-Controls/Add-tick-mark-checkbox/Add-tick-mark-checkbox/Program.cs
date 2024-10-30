using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");

using (WordDocument wordDocument = new WordDocument())
{
    // Add a section and paragraph to initialize an empty document structure
    wordDocument.EnsureMinimal();

    // Append the label "Gender: Female" next to the first checkbox
    wordDocument.LastParagraph.AppendText("Gender:\tFemale ");
    // Append an inline content control to add a checkbox for the checked box
    IInlineContentControl checkInline = wordDocument.LastParagraph.AppendInlineContentControl(ContentControlType.CheckBox);
    // Create a CheckBoxState for the checked state, using a tick symbol in the Wingdings font
    CheckBoxState tickState = new CheckBoxState
    {
        Font = "Wingdings",
        Value = "\uF0FE" // Unicode for the tick symbol (✓) in Wingdings
    };
    // Set the checked state of the checkbox content control to display the tick symbol when selected
    checkInline.ContentControlProperties.CheckedState = tickState;
    // Create a CheckBoxState for the unchecked state, using an empty box symbol in the Wingdings font
    CheckBoxState uncheckedState = new CheckBoxState
    {
        Font = "Wingdings",
        Value = "\uF0A8" // Unicode for the empty box symbol in Wingdings
    };
    // Set the unchecked state of the checkbox content control to display an empty box when not selected
    checkInline.ContentControlProperties.UncheckedState = uncheckedState;
    // Set the initial state of the "Female" checkbox to checked
    checkInline.ContentControlProperties.IsChecked = true;

    // Append a tab space and add the label "Male" for the second checkbox
    wordDocument.LastParagraph.AppendText("\tMale ");
    // Append an inline content control to add a checkbox for the unchecked option
    IInlineContentControl uncheckInline = wordDocument.LastParagraph.AppendInlineContentControl(ContentControlType.CheckBox);
    // Set the checked and unchecked states.
    uncheckInline.ContentControlProperties.CheckedState = tickState;
    uncheckInline.ContentControlProperties.UncheckedState = uncheckedState;
    // Set the initial state of the "Male" checkbox to unchecked
    uncheckInline.ContentControlProperties.IsChecked = false;
    // Save the document to a file in DOCX format
    using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
    {
        wordDocument.Save(outputStream1, FormatType.Docx);
    }
}
