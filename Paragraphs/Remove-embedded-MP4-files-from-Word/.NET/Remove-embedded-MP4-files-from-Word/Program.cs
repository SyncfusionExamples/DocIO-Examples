using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Load the Word document from the specified path
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Find all OLE objects in the document where the ObjectType is "Package" (i.e., embedded file)
    List<Entity> oleObjects = document.FindAllItemsByProperty(EntityType.OleObject, "ObjectType", "Package");

    if (oleObjects != null)
    {
        // Iterate through each OLE object found in the document
        foreach (Entity entity in oleObjects)
        {
            // Cast the entity to a WOleObject
            WOleObject oleObject = entity as WOleObject;

            // Get the native (embedded) data from the OLE object
            byte[] nativeData = oleObject.NativeData;

            // Check if the embedded file is a .mp4 video
            if (oleObject.PackageFileName.EndsWith(".mp4"))
            {
                // Get the paragraph that owns the OLE object
                WParagraph ownerPara = oleObject.OwnerParagraph;

                // Loop to remove all related field code entities until the FieldEnd is reached
                while (true)
                {
                    // Get the next sibling entity in the paragraph
                    Entity nextEntity = entity.NextSibling as Entity;

                    // Remove the sibling entity from the paragraph
                    ownerPara.ChildEntities.Remove(nextEntity);

                    // Stop when the field end marker is reached
                    if ((nextEntity is WFieldMark) && (nextEntity as WFieldMark).Type == FieldMarkType.FieldEnd)
                        break;
                }

                // Finally, remove the OLE object itself from the paragraph
                ownerPara.ChildEntities.Remove(entity);
            }
        }
    }

    // Save the modified Word document to the output path
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}
