using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens an existing Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        //Finds the first occurrence of a particular text in the document.
        TextSelection[] textSelection = document.FindAll("Adventure Works", false, false);
        //Gets the found text as single text range.
        WTextRange textRange = textSelection[0].GetAsOneRange();
        //Get the owner paragraph
        WParagraph ownerPara = textRange.OwnerParagraph;
        //Get the index of text range.
        int index = ownerPara.ChildEntities.IndexOf(textRange);
        //Replace the text.
        document.Replace(textSelection[0].SelectedText, "Adventure Works Cycles", false, false);
        //Create line break.
        Break lineBreak = new Break(document, BreakType.LineBreak);
        //Insert line break in specific index.
        ownerPara.ChildEntities.Insert(index + 1, lineBreak);
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}