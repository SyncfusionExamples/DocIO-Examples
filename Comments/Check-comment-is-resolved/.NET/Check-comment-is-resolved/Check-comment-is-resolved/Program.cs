using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream inputStream = new FileStream(Path.GetFullPath("Data/Sample.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load an existing Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        //Loop through the comments in word document.
        foreach (WComment comment in document.Comments)
        {
            //Check whether the comment is resolved or not.
            Console.WriteLine(comment.Done ? "Resolved" : "Unresolved");
        }
    }
}