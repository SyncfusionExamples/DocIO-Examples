using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Drawing;

using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens an existing Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        // Finds the first occurrence of a particular text that matches the exact case in the document.
        TextSelection textSelection = document.Find("adventure", true, false);
        //Gets the found text as single text range.
        WTextRange textRange = textSelection.GetAsOneRange();
        //Modifies the text.
        textRange.Text = "Replaced text";
        //Sets highlight color.
        textRange.CharacterFormat.HighlightColor = Color.Yellow;
        //Creates file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
