using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System.IO;

namespace Remove_placeholder_of_empty_meta_property
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(docStream, FormatType.Automatic))
                {
                    //Iterate section in the Word document.
                    foreach (WSection section in document.Sections)
                    {
                        //Access the Body of the section where all the contents in the document are apart.
                        WTextBody sectionBody = section.Body;
                        IterateTextBody(sectionBody);
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"Output/Output.docx")) { UseShellExecute = true });
        }
        /// <summary>
        /// Iterate the TextBody of the Word document.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterate through child entities of the WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in the DocIO DOM. 
                //Access the body items (should be either paragraph, table, or block content control) as IEntity.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //Get the element type by using the EntityType.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Iterate through the paragraph's DOM.
                        IterateParagraph(paragraph.Items);
                        break;
                    case EntityType.Table:
                        //Iterate through the table's DOM.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Check whether the content control is xml mapped with meta property.
                        //Also check whether the corresponding meta property value is empty.
                        //If the value is empty, remove the content control.
                        if (IsRemoveContentControl(blockContentControl))
                        {
                            textBody.ChildEntities.Remove(blockContentControl);
                            i--;
                        }
                        break;
                }
            }
        }
        /// <summary>
        /// Iterate Table in the Word document.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterate the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Reuse the code meant for iterating the TextBody.
                    IterateTextBody(cell);
                }
            }
        }
        /// <summary>
        /// Iterate Paragraph in the Word document.
        /// </summary>
        private static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //Get the element type by using EntityType.
                switch (entity.EntityType)
                {
                    case EntityType.TextBox:
                        //Iterate to the body items of the textbox.
                        WTextBox textBox = entity as WTextBox;
                        IterateTextBody(textBox.TextBoxBody);
                        break;
                    case EntityType.Shape:
                        //Iterate to the body items of shape.
                        Shape shape = entity as Shape;
                        IterateTextBody(shape.TextBody);
                        break;
                    case EntityType.InlineContentControl:
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        //Check whether the content control is xml mapped with meta property.
                        //Also check whether the corresponding meta property value is empty.
                        //If the value is empty, remove the content control.
                        if (IsRemoveContentControl(inlineContentControl))
                        {
                            paraItems.Remove(inlineContentControl);
                            i--;
                        }
                        break;
                }
            }
        }
        /// <summary>
        /// Check whether the content control is xml mapped with meta property.
        /// </summary>
        /// <param name="entity">The content control.</param>
        /// <returns>Returns true if content control is needed to remove. Otherwise, false.</returns>
        private static bool IsRemoveContentControl(IEntity entity)
        {
            switch (entity.EntityType)
            {
                case EntityType.BlockContentControl:
                    BlockContentControl blockContentControl = entity as BlockContentControl;
                    ContentControlProperties blockproperties = blockContentControl.ContentControlProperties;
                    if (blockproperties.XmlMapping.IsMapped && !string.IsNullOrEmpty(blockproperties.XmlMapping.XPath)
                        && IsEmptyMetaProperty(blockproperties.Title, entity.Document))
                        return true;
                    break;
                case EntityType.InlineContentControl:
                    InlineContentControl inlineContentControl = entity as InlineContentControl;
                    ContentControlProperties inlineProperties = inlineContentControl.ContentControlProperties;
                    if (inlineProperties.XmlMapping.IsMapped && !string.IsNullOrEmpty(inlineProperties.XmlMapping.XPath)
                        && IsEmptyMetaProperty(inlineProperties.Title, entity.Document))
                        return true;
                    break;
            }
            return false;
        }
        /// <summary>
        /// Check whether the corresponding meta property value is empty.
        /// </summary>
        /// <param name="title">The content control title.</param>
        /// <param name="document">The Word document.</param>
        /// <returns>Returns true if the meta property value is empty. Otherwise, false.</returns>
        private static bool IsEmptyMetaProperty(string title, WordDocument document)
        {
            MetaProperties metaProperties = document.ContentTypeProperties;
            //Iterate through the child entities of metaproperties.
            for (int i = 0; i < metaProperties.Count; i++)
            {
                //Check for a particular display name of metadata and ensure its value is empty or not.
                if (metaProperties[i].DisplayName == title && metaProperties[i].Value == null)
                    return true;
            }
            return false;
        }
    }
}
