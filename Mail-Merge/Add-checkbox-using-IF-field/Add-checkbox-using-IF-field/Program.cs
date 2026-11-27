using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
{
    //Opens the template Word document.
    using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
    {
        // Set default values for the merge fields
        string[] fieldNames = { "Name", "Email", "AgreeToTerms", "SubscribeNewsletter" };
        string[] fieldValues = { "Nancy", "nancy@example.com", "Yes", "No" };

        // Execute mail merge to replace merge fields with actual values.
        wordDocument.MailMerge.Execute(fieldNames, fieldValues);

        // Update any fields in the document to reflect the changes made during the mail merge.
        wordDocument.UpdateDocumentFields();

        // Set up font substitution settings for handling missing fonts.
        wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;

        // Save the modified document.
        using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
        {
            wordDocument.Save(outputStream1, FormatType.Docx);
        }
    }
}

/// <summary>
/// Handles font substitution when a required font is unavailable in the document.
/// </summary>
static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
{
    // Display the name of the original font that needs substitution
    Console.WriteLine(args.OriginalFontName);
}