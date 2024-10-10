using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Register Syncfusion license for the application.
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");


using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
{
    //Load the Word document from the FileStream.
    WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);
    //Iterate through each section in the document.
    for (int secIndex = 0; secIndex < document.Sections.Count; secIndex++)
    {
        WSection sec = document.Sections[secIndex];
        //Iterate through the body items of the section.
        for (int bodyItemIndex = 0; bodyItemIndex < sec.Body.ChildEntities.Count; bodyItemIndex++)
        {
            IEntity body = sec.Body.ChildEntities[bodyItemIndex];
            //Check if the body item is a block content control.
            if (body is BlockContentControl)
            {
                BlockContentControl blockContentControl = body as BlockContentControl;
                //Get the block content control index in the body.
                int index = sec.Body.ChildEntities.IndexOf(blockContentControl);
                //Move the child entities of block content control to section body.
                for (int blockItemIndex = 0; blockItemIndex < blockContentControl.ChildEntities.Count;)
                {
                    IEntity item = blockContentControl.ChildEntities[blockItemIndex];
                    //Insert the child entity to the section body.
                    sec.Body.ChildEntities.Insert(index, item);
                    //Increment the index.
                    index++;
                }
                //Remove the block content control from the section body.
                sec.Body.ChildEntities.Remove(blockContentControl);
            }
            //Check if the body item is a paragraph.
            else if (body is WParagraph)
            {
                WParagraph paragraph = body as WParagraph;
                //Iterate through the items within the paragraph.
                for (int paraItemIndex = 0; paraItemIndex < paragraph.ChildEntities.Count; paraItemIndex++)
                {
                    ParagraphItem item = paragraph.ChildEntities[paraItemIndex] as ParagraphItem;
                    //Check if the paragraph item is an inline content control.
                    if (item is InlineContentControl)
                    {
                        InlineContentControl inlineContentControl = item as InlineContentControl;
                        //Get the index of inline content control in the paragraph.
                        int index = paragraph.ChildEntities.IndexOf(item);
                        //Move the child items of inline content control to the paragraph.
                        for (int inlineItemIndex = 0; inlineItemIndex < inlineContentControl.ParagraphItems.Count;)
                        {
                            ParagraphItem inlineItem = inlineContentControl.ParagraphItems[inlineItemIndex];
                            //Insert the inline content control items to paragraph's child entities.
                            paragraph.ChildEntities.Insert(index, inlineItem);
                            //Increment the index.
                            index++;
                        }
                        //Remove the inline content control from the paragraph.
                        paragraph.ChildEntities.Remove(inlineContentControl);
                    }
                }
            }
        }
    }
    //Create a file stream for saving the document.
    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
    {
        //Save the modified document to the file stream.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
