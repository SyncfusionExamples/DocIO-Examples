using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;


namespace Find_And_Remove_LongText
{
    class Program
    {
        // Set text length to be remove from the document
        static int requiredLongTextLength = 300;
		
        static void Main(string[] args)
        {
            // Open an existing word document 
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(inputFileStream, FormatType.Docx))
                {
                    // Remove long texts
                    CheckAndRemoveLongText(wordDocument);
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            //Saves the PDF file to file system.    
                            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"..\..\..\Output\WordToPDF.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                            {
                                pdfDocument.Save(outputStream);
                            }
                        }
                    }
                }
            }
        }
        private static void CheckAndRemoveLongText(WordDocument document)
        {
            List<Entity> paragraphs = document.FindAllItemsByProperty(EntityType.Paragraph, null, null);
            foreach (Entity paragraph in paragraphs) 
            {
                WParagraph wParagraph = paragraph as WParagraph;
                IterateParagraph(wParagraph.Items);
            }
        }
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
                        WTextRange textRange = entity as WTextRange;
                        //Find and remove the long text in Word document.
                        if (textRange.Text.Length >= requiredLongTextLength)
                        {
                            (entity as WTextRange).Text = string.Empty;
                        }
                        break;
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
                        //Iterates to the paragraph items of inline content control.
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        IterateParagraph(inlineContentControl.ParagraphItems);
                        break;
                }
            }
        }
    }
}
