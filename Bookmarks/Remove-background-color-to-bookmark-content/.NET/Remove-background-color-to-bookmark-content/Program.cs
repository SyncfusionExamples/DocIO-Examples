using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Remove_background_color_to_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an input file stream to open the document
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(inputStream,FormatType.Docx))
                {
                    // Create the bookmark navigator instance
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    // Move to the bookmark
                    bookmarkNavigator.MoveToBookmark("Adventure_Bkmk");
                    // Get the bookmark content
                    TextBodyPart part = bookmarkNavigator.GetBookmarkContent();
                    // Iterate through the content (implement this according to your needs)
                    IterateTextBodyPart(part.BodyItems);
                    // Replace the bookmark content with modified content
                    bookmarkNavigator.ReplaceBookmarkContent(part);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            } 
        }
        /// <summary>
        /// Iterates through the body contents of the Word document.
        /// </summary>
        /// <param name="textBodyItems"></param>
        private static void IterateTextBodyPart(EntityCollection textBodyItems)
        {
            // Iterates through each of the child items of WTextBody
            for (int i = 0; i < textBodyItems.Count; i++)
            {
                // IEntity is the basic unit in DocIO DOM.
                // Accesses the body items (should be either paragraph, table or block content control) as IEntity
                IEntity bodyItemEntity = textBodyItems[i];

                // A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                // Decide the element type by using EntityType
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        {
                            WParagraph paragraph = bodyItemEntity as WParagraph;
                            if (paragraph != null)
                            {
                                // Processes the paragraph contents
                                // Iterates through the paragraph's DOM
                                IterateParagraph(paragraph.Items);
                            }
                            break;
                        }
                    case EntityType.Table:
                        {
                            // Table is a collection of rows and cells
                            // Iterates through table's DOM
                            WTable table = bodyItemEntity as WTable;
                            if (table != null)
                            {
                                IterateTable(table);
                            }
                            break;
                        }
                    case EntityType.BlockContentControl:
                        {
                            BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                            if (blockContentControl != null && blockContentControl.TextBody != null)
                            {
                                // Iterates the body items of Block Content Control.
                                IterateTextBodyPart(blockContentControl.TextBody.ChildEntities);
                            }
                            break;
                        }
                }
            }
        }
        /// <summary>
        /// Iterates through the table.
        /// </summary>
        /// <param name="table"></param>
        private static void IterateTable(WTable table)
        {
            //Iterates the row collection in a table
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row
                foreach (WTableCell cell in row.Cells)
                {
                    //Table cell is derived from (also a) TextBody
                    //Reusing the code meant for iterating TextBody
                    IterateTextBodyPart(cell.ChildEntities);
                }
            }
        }
        /// <summary>
        /// Iterates through the paragraph.
        /// </summary>
        /// <param name="paraItems"></param>
        private static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //A paragraph can have child elements such as text, image, hyperlink, symbols, etc.,
                //Decides the element type by using EntityType
                switch (entity.EntityType)
                {
                    case EntityType.TextRange:
                        //Remove the text back color
                        WTextRange textRange = entity as WTextRange;
                        textRange.CharacterFormat.TextBackgroundColor = Syncfusion.Drawing.Color.Empty;
                        break;              
                    case EntityType.TextBox:
                        //Iterates to the body items of textbox.
                        WTextBox textBox = entity as WTextBox;
                        IterateTextBodyPart(textBox.TextBoxBody.ChildEntities);
                        break;
                    case EntityType.Shape:
                        //Iterates to the body items of shape.
                        Shape shape = entity as Shape;
                        IterateTextBodyPart(shape.TextBody.ChildEntities);
                        break;
                    case EntityType.InlineContentControl:
                        //Iterates to the paragraph items of inline content control.
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        IterateParagraph(inlineContentControl.ParagraphItems);
                        break;
                }
            }
        }
    }
}
