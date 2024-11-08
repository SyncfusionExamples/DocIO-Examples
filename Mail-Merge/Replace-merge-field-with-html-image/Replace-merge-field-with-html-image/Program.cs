using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Data;

//Dictionary to maintain paragraph and corresponding merge field index with HTML content.
Dictionary<WParagraph, Dictionary<int, string>> paraToInsertHTML = new Dictionary<WParagraph, Dictionary<int, string>>();

using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Opens the template Word document.
    using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
    {
        //Creates the mail merge events handler to replace merge field with HTML.
        document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
        //Gets data to perform the mail merge.
        DataTable table = GetDataTable();
        //Performs the mail merge.
        document.MailMerge.Execute(table);
        //Append HTML to paragraph.
        InsertHtml();
        //Removes the mail merge events handler.
        document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Modified Word document.
            document.Save(outputStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Handles the MergeField event to replace the merge field with HTML content.
/// </summary>
void MergeFieldEvent(object sender, MergeFieldEventArgs args)
{
    if (args.FieldName.Equals("Logo"))
    {
        //Gets the current merge field owner paragraph.
        WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
        //Gets the current merge field index in the current paragraph.
        int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
        //Maintain HTML in collection.
        Dictionary<int, string> fieldValues = new Dictionary<int, string>();
        fieldValues.Add(mergeFieldIndex, args.FieldValue.ToString());
        //Maintain paragraph in collection.
        paraToInsertHTML.Add(paragraph, fieldValues);
        //Set field value as empty.
        args.Text = string.Empty;
    }
}

/// <summary>
/// Get a data table for the mail merge operation.
/// </summary>
DataTable GetDataTable()
{
    DataTable dataTable = new DataTable("HTML");
    dataTable.Columns.Add("CustomerName");
    dataTable.Columns.Add("Address");
    dataTable.Columns.Add("Phone");
    dataTable.Columns.Add("Logo");

    //Adds sample data to the DataTable.
    DataRow datarow = dataTable.NewRow();
    dataTable.Rows.Add(datarow);
    datarow["CustomerName"] = "Nancy Davolio";
    datarow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
    datarow["Phone"] = "1-888-936-8638";

    //Reads HTML string from the file.
    string htmlString = File.ReadAllText(Path.GetFullPath(@"Data/File.html"));
    datarow["Logo"] = htmlString;

    return dataTable;
}

/// <summary>
/// Inserts HTML content into the specified paragraphs and positions within the Word document.
/// </summary>
void InsertHtml()
{
    //Iterates through each item in the dictionary.
    foreach (KeyValuePair<WParagraph, Dictionary<int, string>> dictionaryItems in paraToInsertHTML)
    {
        WParagraph paragraph = dictionaryItems.Key as WParagraph;
        Dictionary<int, string> values = dictionaryItems.Value as Dictionary<int, string>;
        //Iterates through each value in the dictionary.
        foreach (KeyValuePair<int, string> valuePair in values)
        {
            int index = valuePair.Key;
            string fieldValue = valuePair.Value;

            //Subscribe the ImageNodeVisited event to resolve images within HTML content.
            paragraph.Document.HTMLImportSettings.ImageNodeVisited += OpenImage;

            //Inserts an HTML string at the same position of mergefield in a Word document.
            paragraph.OwnerTextBody.InsertXHTML(fieldValue, paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph), index);

            //Unsubscribe the ImageNodeVisited event after processing.
            paragraph.Document.HTMLImportSettings.ImageNodeVisited -= OpenImage;
        }
    }
    //Clears the dictionary after inserting HTML content.
    paraToInsertHTML.Clear();
}

/// <summary>
/// Opens images referenced within HTML content.
/// </summary>
void OpenImage(object sender, ImageNodeVisitedEventArgs args)
{
    //Reads the image from the specified URI path and assigns it to the image stream.
    args.ImageStream = File.OpenRead(args.Uri);
}
