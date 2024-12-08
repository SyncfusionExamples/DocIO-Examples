using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

//Open the file as stream.
using (FileStream inputDocumentStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Create a new Word document.
    using (WordDocument document = new WordDocument(inputDocumentStream, FormatType.Docx))
    {
        //Find drop-down content control by title.
        InlineContentControl inlineContentControl = document.FindItemByProperty(EntityType.InlineContentControl, "ContentControlProperties.Title", "Status") as InlineContentControl;
        WTextRange textRange = inlineContentControl.ParagraphItems[0] as WTextRange;
        //Select drop-down
        textRange.Text = inlineContentControl.ContentControlProperties.ContentControlListItems[1].DisplayText;

        //Find date content control by tag.
        inlineContentControl = document.FindItemByProperty(EntityType.InlineContentControl, "ContentControlProperties.Tag", "Date") as InlineContentControl;
        textRange = inlineContentControl.ParagraphItems[0] as WTextRange;
        //Set today's date to display.
        textRange.Text = DateTime.Now.ToShortDateString();

        //Find text content control by title.
        inlineContentControl = document.FindItemByProperty(EntityType.InlineContentControl, "ContentControlProperties.Title", "ProjectName") as InlineContentControl;
        //Fill text.
        textRange = inlineContentControl.ParagraphItems[0] as WTextRange;
        textRange.Text = "Website for Adventure works cycle";

        //Find checkbox content control by type.
        inlineContentControl = document.FindItemByProperty(EntityType.InlineContentControl, "ContentControlProperties.Type", "CheckBox") as InlineContentControl;
        //Check the checkbox
        inlineContentControl.ContentControlProperties.IsChecked = true;

        //Create file stream.
        using (FileStream outputDocumentStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Save the Word document to file stream.
            document.Save(outputDocumentStream, FormatType.Docx);
        }
    }
}
