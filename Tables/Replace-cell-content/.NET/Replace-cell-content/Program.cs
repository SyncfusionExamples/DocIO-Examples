using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputFileStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Open the template Word document.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        //Retrieve the first section of the document.
        IWSection section = document.LastSection;
        //Get the first table in the section.
        WTable table = section.Body.Tables[0] as WTable;
        //Access the specific cells by their indices.
        WTableCell cell1 = table[1, 1];
        WTableCell cell2 = table[2, 2];
        //Clear the contents of the first cell.
        cell1.ChildEntities.Clear();
        //Add a new paragraph with content to the first cell.
        cell1.AddParagraph().AppendText("Adventure");
        //Clear the contents of the second cell.
        cell2.ChildEntities.Clear();
        //Add a new paragraph with content to the second cell.
        cell2.AddParagraph().AppendText("Cycle");
        //Save the modified document.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
        {
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}