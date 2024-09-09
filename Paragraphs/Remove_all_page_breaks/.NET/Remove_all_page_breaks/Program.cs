using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    //Open the existing Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        //Get the list of page breaks
        List<Entity> entities = document.FindAllItemsByProperty(EntityType.Break, "BreakType", "PageBreak");

        //Iterate through all page breaks and remove it from the owner paragraph
        foreach (Entity entity in entities)
            (entity as Break).OwnerParagraph.ChildEntities.Remove(entity);

        //Save the Word document
        using (FileStream outputStream = new FileStream(@"Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}