using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_a_Checkbox_in_a_Word_Document
{
    class Program
    {

        static void Main(string[] args)
        {
            // Load the word document
            using (FileStream fileStream = new FileStream("../../../Data/Template.docx", FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    foreach (WSection section in document.Sections)
                    {
                        //Accesses the Body of section where all the contents in document are apart
                        WTextBody sectionBody = section.Body;
                        IterateTextBody(sectionBody);
                        WHeadersFooters headersFooters = section.HeadersFooters;
                        //Consider that OddHeader and OddFooter are applied to this document
                        //Iterates through the TextBody of OddHeader and OddFooter
                        IterateTextBody(headersFooters.OddHeader);
                        IterateTextBody(headersFooters.OddFooter);
                    }
                    using (FileStream outputStream = new FileStream("../../../Output/Result.docx", FileMode.Create, FileAccess.Write))
                    {
                        //Saves the stream as Word file
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        public static void IterateTextBody(WTextBody textBody)
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
                        if (blockContentControl.ContentControlProperties.Type == ContentControlType.CheckBox)
                            blockContentControl.ContentControlProperties.IsChecked = false;
                        //Iterates to the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }
        public static void IterateTable(WTable table)
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
        public static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //A paragraph can have child elements such as text, image, hyperlink, symbols, etc.,
                //Decides the element type by using EntityType
                switch (entity.EntityType)
                {
                    case EntityType.CheckBox:
                        if (entity is WCheckBox)
                            (entity as WCheckBox).Checked = false;
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
                        if (inlineContentControl.ContentControlProperties.Type == ContentControlType.CheckBox)
                            inlineContentControl.ContentControlProperties.IsChecked = false;
                        IterateParagraph(inlineContentControl.ParagraphItems);
                        break;
                }
            }
        }
    }
}
