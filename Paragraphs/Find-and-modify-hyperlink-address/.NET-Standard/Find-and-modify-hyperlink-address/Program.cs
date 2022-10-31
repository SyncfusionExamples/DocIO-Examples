using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Find_and_modify_hyperlink_address
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access paragraph in a Word document.
                    WParagraph paragraph = document.Sections[0].Paragraphs[1];
                    WField field = paragraph.ChildEntities[0] as WField;
                    //Create an instance of hyperlink.
                    Hyperlink hyperlink = new Hyperlink(field);
                    //Set the hyperlink type, URL and the text to display.
                    hyperlink.Type = HyperlinkType.WebLink;
                    hyperlink.Uri = "http://www.google.com";
                    hyperlink.TextToDisplay = "Google";
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
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
