using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Edit_text_in_inline_content_control
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Creates an instance of WordDocument class
                using (WordDocument document = new WordDocument(docStream, FormatType.Automatic))
                {
                    ///Processes the body contents for each section in the Word document
                    foreach (WSection section in document.Sections)
                    {
                        //Accesses the Body of section where all the contents in document are apart
                        WTextBody sectionBody = section.Body;
                        IterateTextBody(sectionBody);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Iterate TextBody
        /// </summary>
        /// <param name="textBody"></param>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterates through each of the child items of WTextBody
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items (should be either paragraph, table or block content control) as IEntity
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                //Decides the element type by using EntityType
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Processes the paragraph contents
                        //Iterates through the paragraph's DOM
                        IterateParagraph(paragraph.Items);
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells
                        //Iterates through table's DOM
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterates to the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }
        /// <summary>
        /// Iterate Table
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
                    IterateTextBody(cell);
                }
            }
        }
        /// <summary>
        /// Iterate Paragraph
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
                    case EntityType.TextBox:
                        //Iterates to the body items of textbox.
                        WTextBox textBox = entity as WTextBox;
                        IterateTextBody(textBox.TextBoxBody);
                        break;
                    case EntityType.Shape:
                        //Iterates to the body items of shape.
                        Shape shape = entity as Shape;
                        IterateTextBody(shape.TextBody);
                        break;
                    case EntityType.InlineContentControl:
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        if((inlineContentControl.ContentControlProperties.Type == ContentControlType.RichText 
                            || inlineContentControl.ContentControlProperties.Type == ContentControlType.Text) 
                            && inlineContentControl.ContentControlProperties.Title == "ReplaceText")
                            ReplaceTextWithInlineContentControl("Hello World", inlineContentControl);
                        break;
                }
            }
        }
        /// <summary>
        /// Replace Text With Inline content control
        /// </summary>
        /// <param name="text"></param>
        /// <param name="inlineContentControl"></param>
        private static void ReplaceTextWithInlineContentControl(string replacementText, InlineContentControl inlineContentControl)
        {
            WCharacterFormat characterFormat = null;
            foreach (ParagraphItem item in inlineContentControl.ParagraphItems)
            {
                if (item is WTextRange)
                {
                    characterFormat = (item as WTextRange).CharacterFormat;
                    break;
                }
            }
            //Remove exiting items and add new text range with required text
            inlineContentControl.ParagraphItems.Clear();
            WTextRange textRange = new WTextRange(inlineContentControl.Document);
            textRange.Text = replacementText;
            if (characterFormat != null)
                textRange.ApplyCharacterFormat(characterFormat);
            inlineContentControl.ParagraphItems.Add(textRange);
        }
    }
}
