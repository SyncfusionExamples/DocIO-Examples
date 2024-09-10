using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Find_an_image_by_title
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Gets textbody content.
                    WTextBody textBody = document.Sections[0].Body;
                    //Retrieves and modify the image based on its title by iterating from the document elements.
                    IterateTextBody(textBody, "Product");
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods

        /// <summary>
        /// Iterates through the textbody.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody, string pictureTitle)
        {
            //Iterates through each of the child items of WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items (should be either paragraph, table or block content control) as IEntity.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                //Decides the element type by using EntityType.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Processes the paragraph contents.
                        //Iterates through the paragraph's DOM.
                        IterateParagraph(paragraph.Items, pictureTitle);
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells.
                        //Iterates through table's DOM.
                        IterateTable(bodyItemEntity as WTable, pictureTitle);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterates to the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody, pictureTitle);
                        break;
                }
            }
        }

        /// <summary>
        /// Iterates through the table.
        /// </summary>
        private static void IterateTable(WTable table, string pictureTitle)
        {
            //Iterates the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Table cell is derived from (also a) TextBody.
                    //Reusing the code meant for iterating TextBody.
                    IterateTextBody(cell, pictureTitle);
                }
            }
        }

        /// <summary>
        /// Iterates through the paragraph.
        /// </summary>
        private static void IterateParagraph(ParagraphItemCollection paraItems, string pictureTitle)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //A paragraph can have child elements such as text, image, hyperlink, symbols, etc.,
                //Decides the element type by using EntityType.
                switch (entity.EntityType)
                {
                    case EntityType.Picture:
                        WPicture picture = entity as WPicture;
                        //Gets the image from its title and modifies its width and height.
                        if (picture.Title == pictureTitle)
                        {
                            picture.Width = 150;
                            picture.Height = 100;
                        }
                        break;
                    case EntityType.TextBox:
                        //Iterates to the body items of textbox.
                        WTextBox textBox = entity as WTextBox;
                        IterateTextBody(textBox.TextBoxBody, pictureTitle);
                        break;
                    case EntityType.Shape:
                        //Iterates to the body items of shape.
                        Shape shape = entity as Shape;
                        IterateTextBody(shape.TextBody, pictureTitle);
                        break;
                    case EntityType.InlineContentControl:
                        //Iterates to the paragraph items of inline content control.
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        IterateParagraph(inlineContentControl.ParagraphItems, pictureTitle);
                        break;
                }
            }
        }
        #endregion
    }
}
