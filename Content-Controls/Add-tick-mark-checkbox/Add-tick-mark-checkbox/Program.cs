using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");

using (WordDocument wordDocument = new WordDocument())
{
    // Add a section & a paragraph in the empty document 
    wordDocument.EnsureMinimal();
    // Append an inline content control for the tick checkbox
    IInlineContentControl tickInline = wordDocument.LastParagraph.AppendInlineContentControl(ContentControlType.CheckBox);
    // Create a CheckBoxState for the tick symbol
    CheckBoxState tickState = new CheckBoxState
    {
        Font = "Wingdings",
        Value = "\uF0FE" // Unicode for tick symbol (✓) in Wingdings
    };
    // Set the checked state of the tick checkbox content control
    tickInline.ContentControlProperties.CheckedState = tickState;
    // Create a CheckBoxState for the unchecked state (empty box)
    CheckBoxState uncheckedState = new CheckBoxState
    {
        Font = "Wingdings",
        Value = "\uF0A8" // Unicode for empty box in Wingdings
    };
    // Set the unchecked state of the tick checkbox content control
    tickInline.ContentControlProperties.UncheckedState = uncheckedState;
    // Set the initial state of the tick checkbox as checked
    tickInline.ContentControlProperties.IsChecked = true;
    using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
    {
        wordDocument.Save(outputStream1, FormatType.Docx);
    }
}