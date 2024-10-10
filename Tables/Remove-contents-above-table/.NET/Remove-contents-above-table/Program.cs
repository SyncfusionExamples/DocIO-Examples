using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Opens the input Word document template from the specified path.
using (FileStream inputFileStream = new FileStream(Path.GetFullPath("Data/Input.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Loads the Word document into a WordDocument object.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        //Finds the table in the document by using the table's Title property, with the value "Table1".
        WTable table = document.FindItemByProperty(EntityType.Table, "Title", "Table1") as WTable;
        //Removes the content that exists before the located table in the document.
        RemoveContentBeforeTable(document, table);
        //Saves the modified document to the specified output path.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
        {
            document.Save(outputFileStream, FormatType.Docx); //Saves the document in DOCX format.
        }
        //Closes the Word document after saving.
        document.Close();
    }
}

/// <summary>
/// Removes all content before the specified table, including content in previous sections, in a Word document.
/// </summary>
void RemoveContentBeforeTable(WordDocument document, WTable inputTable)
{
    //Get the index of the input table.
    int tableIndex = inputTable.OwnerTextBody.ChildEntities.IndexOf(inputTable);
    //Get the section entity.
    WSection currSection = GetOwnerSection(inputTable, ref tableIndex);
    //Get the section index.
    int sectionIndex = document.Sections.IndexOf(currSection);
    //Remove the items before the table in the current section.
    for (int i = tableIndex - 1; i >= 0; i--)
        currSection.Body.ChildEntities.RemoveAt(i);
    //Remove the previous sections.
    for (int i = sectionIndex - 1; i >= 0; i--)
        document.Sections.RemoveAt(i);
}

/// <summary>
/// Traverses the entity hierarchy to find its owning section and updates the table index accordingly.
/// </summary>
WSection GetOwnerSection(Entity entity, ref int tableIndex)
{
    while (!(entity is WSection))
    {
        //If the entity is table, then get the table index.
        if (entity is WTable)
        {
            WTable table = entity as WTable;
            tableIndex = table.OwnerTextBody.ChildEntities.IndexOf(table);
        }
        //If the entity is block content control, then remove the child entities before the table.
        //and consider the block content control index as tableIndex.
        else if (entity is BlockContentControl)
        {
            BlockContentControl blockContentControl = entity as BlockContentControl;
            //Remove the child entitites of block content control before the table.
            for (int i = tableIndex - 1; i >= 0; i--)
                blockContentControl.TextBody.ChildEntities.RemoveAt(i);
            //Get the block content control index.
            tableIndex = blockContentControl.OwnerTextBody.ChildEntities.IndexOf(blockContentControl);
        }
        //Move to the owner entity.
        entity = entity.Owner;
    }
    return entity as WSection;
}