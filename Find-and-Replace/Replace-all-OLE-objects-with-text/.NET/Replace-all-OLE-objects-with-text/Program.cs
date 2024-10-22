using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open the file as a Stream.
using (FileStream docStream = new FileStream("Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    //Load the file stream into a Word document.
    using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
    {
        //Find all OLE object by EntityType in Word document.
        List<Entity> oleObjects = document.FindAllItemsByProperty(EntityType.OleObject, null, null);
        //Remove the OLE objects and endnotes.
        for (int i = 0; i < oleObjects.Count; i++)
        {
            WOleObject ole = oleObjects[i] as WOleObject;
            //Replaces the OLE object with a alternate text.
            ReplaceOLEObjectsWithPlaceHolder(ole, "Embedded file was here");
        }
        //Save a  Word document to the MemoryStream.
        FileStream outputStream = new FileStream(@"Output/Output.docx", FileMode.OpenOrCreate);
        document.Save(outputStream, FormatType.Docx);
        //Closes the Word document
        document.Close();
        outputStream.Close();
    }
}

void ReplaceOLEObjectsWithPlaceHolder(WOleObject ole, string replacingText)
{
    WParagraph ownerPara = ole.OwnerParagraph;
    int index = ownerPara.ChildEntities.IndexOf(ole);
    //Removes the ole object.
    RemoveOLEObject(ownerPara, index);
    //Insert the alternate text.
    InsertTextrange(ownerPara, index, replacingText);
}

void RemoveOLEObject(WParagraph ownerPara, int index)
{
    //Iterate from FieldEnd to OLE object
    for (int i = index + 4; i >= index; i--)
    {
        //Remove the OLE object based on the structure
        ownerPara.ChildEntities.RemoveAt(i);
    }
}

void InsertTextrange(WParagraph ownerPara, int index, string replacingText)
{
    //Create a new textrange
    WTextRange textRange = new WTextRange(ownerPara.Document);
    //Add the text
    textRange.Text = replacingText;
    //Insert the textrange in the particular index
    ownerPara.ChildEntities.Insert(index, textRange);
}